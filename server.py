"""
AutoCAD MCP Server
==================
A Model Context Protocol server that lets Claude (or any MCP client) drive
AutoCAD on Windows through the AutoCAD ActiveX/COM automation interface.

Architecture:
    MCP client  <-- stdio JSON-RPC -->  this server  <-- COM -->  AutoCAD

Requires:
    - Windows
    - AutoCAD 2021 or newer, installed and licensed
    - Python 3.10+
    - pywin32   (provides the win32com COM bridge)
    - mcp[cli]  (the official MCP Python SDK / FastMCP)
"""

import logging
import sys

from mcp.server.fastmcp import FastMCP

# IMPORTANT (stdio servers): never print() to stdout, it corrupts the
# JSON-RPC stream. Log to stderr instead.
logging.basicConfig(level=logging.INFO, stream=sys.stderr)
log = logging.getLogger("autocad-mcp")

mcp = FastMCP("autocad")

# ---------------------------------------------------------------------------
# AutoCAD connection (lazy + cached)
# ---------------------------------------------------------------------------
_acad = None  # cached AutoCAD COM object

# AutoCAD registers a version-specific COM ProgID (e.g. "AutoCAD.Application.26"
# for AutoCAD 2027). The generic "AutoCAD.Application" isn't always present, so we
# try the generic name first and then fall back across recent version numbers.
# ZWCAD (an AutoCAD-compatible CAD app) exposes the same object model under the
# ProgID "ZWCAD.Application", so we try it too -- the drawing tools below work the
# same way against either program.
_PROGIDS = [
    "AutoCAD.Application",
    "AutoCAD.Application.26",
    "AutoCAD.Application.25",
    "AutoCAD.Application.24",
    "AutoCAD.Application.23",
    "ZWCAD.Application",
]


def _get_acad():
    """Return a live AutoCAD COM object, attaching to a running AutoCAD or launching one."""
    global _acad
    # win32com is imported lazily so the module can still be inspected on
    # non-Windows machines (e.g. when reading the code on GitHub).
    import win32com.client

    if _acad is not None:
        # Verify the cached connection is still alive. If the CAD program was
        # closed since we last connected, the reference goes stale and any call
        # raises "RPC server unavailable" -- so drop it and reconnect below.
        try:
            _ = _acad.Name
            return _acad
        except Exception:
            _acad = None

    last_error = None
    # First, try to attach to an AutoCAD that is already open.
    for progid in _PROGIDS:
        try:
            _acad = win32com.client.GetActiveObject(progid)
            _acad.Visible = True
            return _acad
        except Exception as e:
            last_error = e
    # Otherwise, launch a new AutoCAD instance.
    for progid in _PROGIDS:
        try:
            _acad = win32com.client.Dispatch(progid)
            _acad.Visible = True
            return _acad
        except Exception as e:
            last_error = e

    raise RuntimeError(
        "Could not connect to AutoCAD or ZWCAD. Make sure the program is open, "
        f"then try again. (Tried ProgIDs {_PROGIDS}; last error: {last_error})"
    )


def _model_space():
    """Return the ModelSpace of the active drawing (creates a doc if none open)."""
    acad = _get_acad()
    if acad.Documents.Count == 0:
        acad.Documents.Add()
    return acad.ActiveDocument.ModelSpace


def _point(x, y, z=0.0):
    """Build a VARIANT array of 3 doubles, the form AutoCAD COM expects for a point."""
    import pythoncom
    import win32com.client
    return win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(x), float(y), float(z)]
    )


def _coords(points):
    """Flatten [(x, y), ...] into a VARIANT array of doubles for polylines."""
    import pythoncom
    import win32com.client
    flat = []
    for x, y in points:
        flat.extend([float(x), float(y)])
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, flat)


# Standard AutoCAD/ZWCAD lineweights, in hundredths of a millimetre.
_LINEWEIGHTS = [0, 5, 9, 13, 15, 18, 20, 25, 30, 35, 40, 50, 53, 60, 70,
                80, 90, 100, 106, 120, 140, 158, 200, 211]


def _nearest_lineweight(mm):
    """Map a thickness in millimetres to the nearest valid lineweight value (1/100 mm)."""
    target = int(round(float(mm) * 100))
    return min(_LINEWEIGHTS, key=lambda v: abs(v - target))


def _show_lineweights():
    """Turn on lineweight display so thickness is visible in model space."""
    try:
        _get_acad().ActiveDocument.SetVariable("LWDISPLAY", 1)
    except Exception:
        pass


# Text justification codes -- the AutoCAD/ZWCAD acAlignment enum. The exact
# integers matter: getting them wrong silently anchors text by the wrong point.
# Reference enum:
#   0 left          1 center         2 right         3 aligned
#   4 middle        5 fit
#   6 top-left      7 top-center     8 top-right
#   9 middle-left  10 middle-center 11 middle-right
#  12 bottom-left  13 bottom-center 14 bottom-right
# "center" is treated as middle-center (10) -- what you want to centre a label
# both horizontally and vertically inside a box.
_ALIGN_CODES = {
    "left": 0,
    "center": 10, "centre": 10, "centered": 10,
    "middlecenter": 10, "middlecentre": 10,
    "right": 2, "aligned": 3, "middle": 4, "fit": 5,
    "topleft": 6, "topcenter": 7, "topright": 8,
    "middleleft": 9, "middleright": 11,
    "bottomleft": 12, "bottomcenter": 13, "bottomright": 14,
}


def _align_code(align):
    return _ALIGN_CODES.get(str(align).lower().replace("_", "").replace(" ", ""), 0)


def _add_text(ms, text, x, y, height, align="left", rotation=0.0):
    """Create a text object, applying justification and rotation if given."""
    import math
    txt = ms.AddText(str(text), _point(x, y), float(height))
    code = _align_code(align)
    if code != 0:
        try:
            txt.Alignment = code
            txt.TextAlignmentPoint = _point(x, y)
        except Exception:
            pass
    if rotation:
        try:
            txt.Rotation = math.radians(float(rotation))
        except Exception:
            pass
    return txt


def _add_arc(ms, cx, cy, radius, start_deg, end_deg):
    """Add an arc, sweeping counter-clockwise from start_deg to end_deg."""
    import math
    return ms.AddArc(_point(cx, cy), float(radius),
                     math.radians(float(start_deg)), math.radians(float(end_deg)))


def _add_rounded_rect(ms, x1, y1, x2, y2, radius):
    """Add a rounded rectangle as four straight sides + four corner arcs.
    Corners are clamped so the radius never exceeds half the shorter side."""
    x1, y1, x2, y2 = float(x1), float(y1), float(x2), float(y2)
    if x2 < x1:
        x1, x2 = x2, x1
    if y2 < y1:
        y1, y2 = y2, y1
    r = max(0.0, min(float(radius), (x2 - x1) / 2.0, (y2 - y1) / 2.0))
    if r == 0:
        poly = ms.AddLightWeightPolyline(
            _coords([(x1, y1), (x2, y1), (x2, y2), (x1, y2)]))
        poly.Closed = True
        return
    # straight sides
    ms.AddLine(_point(x1 + r, y1), _point(x2 - r, y1))  # bottom
    ms.AddLine(_point(x2, y1 + r), _point(x2, y2 - r))  # right
    ms.AddLine(_point(x2 - r, y2), _point(x1 + r, y2))  # top
    ms.AddLine(_point(x1, y2 - r), _point(x1, y1 + r))  # left
    # corner arcs (CCW)
    _add_arc(ms, x1 + r, y1 + r, r, 180, 270)  # bottom-left
    _add_arc(ms, x2 - r, y1 + r, r, 270, 360)  # bottom-right
    _add_arc(ms, x2 - r, y2 - r, r, 0, 90)     # top-right
    _add_arc(ms, x1 + r, y2 - r, r, 90, 180)   # top-left


# ---------------------------------------------------------------------------
# Tools
# ---------------------------------------------------------------------------
@mcp.tool()
def connect() -> str:
    """Connect to the CAD application (AutoCAD or ZWCAD), launching it if needed, and report which one."""
    acad = _get_acad()
    try:
        app_name = acad.Name  # "AutoCAD" or "ZWCAD"
    except Exception:
        app_name = "the CAD application"
    try:
        doc_name = acad.ActiveDocument.Name
    except Exception:
        acad.Documents.Add()
        doc_name = acad.ActiveDocument.Name
    return f"Connected to {app_name}. Active drawing: {doc_name}"


@mcp.tool()
def draw_line(x1: float, y1: float, x2: float, y2: float, lineweight_mm: float = 0.0) -> str:
    """Draw a line in model space from (x1, y1) to (x2, y2). Optionally set its thickness in mm
    (lineweight_mm); leave 0 to use the layer's default thickness."""
    ms = _model_space()
    line = ms.AddLine(_point(x1, y1), _point(x2, y2))
    if lineweight_mm and lineweight_mm > 0:
        line.Lineweight = _nearest_lineweight(lineweight_mm)
        _show_lineweights()
    return f"Line drawn from ({x1}, {y1}) to ({x2}, {y2})."


@mcp.tool()
def draw_circle(center_x: float, center_y: float, radius: float) -> str:
    """Draw a circle in model space at (center_x, center_y) with the given radius."""
    ms = _model_space()
    ms.AddCircle(_point(center_x, center_y), float(radius))
    return f"Circle drawn at ({center_x}, {center_y}), radius {radius}."


@mcp.tool()
def draw_rectangle(x1: float, y1: float, x2: float, y2: float) -> str:
    """Draw a closed rectangle from corner (x1, y1) to opposite corner (x2, y2)."""
    ms = _model_space()
    pts = [(x1, y1), (x2, y1), (x2, y2), (x1, y2)]
    poly = ms.AddLightWeightPolyline(_coords(pts))
    poly.Closed = True
    return f"Rectangle drawn from ({x1}, {y1}) to ({x2}, {y2})."


@mcp.tool()
def draw_polyline(points: list[list[float]], lineweight_mm: float = 0.0) -> str:
    """Draw a polyline through a list of [x, y] points, e.g. [[0,0],[10,0],[10,10]]. Optionally
    set its thickness in mm (lineweight_mm); leave 0 to use the layer's default thickness."""
    if len(points) < 2:
        return "Need at least 2 points to draw a polyline."
    ms = _model_space()
    poly = ms.AddLightWeightPolyline(_coords([(p[0], p[1]) for p in points]))
    if lineweight_mm and lineweight_mm > 0:
        poly.Lineweight = _nearest_lineweight(lineweight_mm)
        _show_lineweights()
    return f"Polyline drawn through {len(points)} points."


@mcp.tool()
def add_text(text: str, x: float, y: float, height: float = 2.5, align: str = "left",
             rotation: float = 0.0) -> str:
    """Place single-line text at (x, y) with the given height. `align` sets justification:
    "left" (default) puts the lower-left of the text at (x, y); "center" centres the text on
    (x, y) -- use that with a box's centre point to centre a label inside it. Other options:
    "middleleft", "topcenter", "bottomcenter", "right", etc. `rotation` is in degrees (90 = text
    running up a vertical wire)."""
    ms = _model_space()
    _add_text(ms, text, x, y, height, align, rotation)
    return f"Text '{text}' placed at ({x}, {y}) [{align}]."


@mcp.tool()
def draw_batch(items: list[dict], layer: str = "") -> str:
    """Draw many objects in ONE call (instead of dozens of separate calls). `items` is a list of
    dicts, each with a "type" plus its parameters:
      {"type": "rectangle", "x1":.., "y1":.., "x2":.., "y2":..}
      {"type": "rounded_rectangle", "x1":.., "y1":.., "x2":.., "y2":.., "radius":..}
      {"type": "line",      "x1":.., "y1":.., "x2":.., "y2":..}
      {"type": "polyline",  "points": [[x,y], [x,y], ...]}
      {"type": "circle",    "center_x":.., "center_y":.., "radius":..}
      {"type": "arc",       "center_x":.., "center_y":.., "radius":.., "start_deg":.., "end_deg":..}
      {"type": "text",      "text":"..", "x":.., "y":.., "height":2.5, "align":"left", "rotation":0}
      {"type": "block",     "name":"3RT2015-1AP01", "x":.., "y":.., "scale":1, "rotation":0,
                            "layer":"", "attributes":{"TAG":"-K1"}}
    The "block" type places a symbol from the drawing's library -- ideal for schematics. Optionally
    pass a top-level `layer` to set the active layer first. Each item is drawn independently, so one
    bad item won't stop the rest; the result reports how many succeeded and any errors."""
    import math
    if layer:
        try:
            acad = _get_acad()
            doc = acad.ActiveDocument
            doc.ActiveLayer = doc.Layers.Item(layer)
        except Exception:
            pass
    ms = _model_space()
    ok = 0
    errors = []
    for i, it in enumerate(items):
        try:
            t = str(it.get("type", "")).lower()
            if t in ("rectangle", "rect", "box"):
                x1, y1, x2, y2 = float(it["x1"]), float(it["y1"]), float(it["x2"]), float(it["y2"])
                poly = ms.AddLightWeightPolyline(
                    _coords([(x1, y1), (x2, y1), (x2, y2), (x1, y2)])
                )
                poly.Closed = True
            elif t == "line":
                ms.AddLine(
                    _point(float(it["x1"]), float(it["y1"])),
                    _point(float(it["x2"]), float(it["y2"])),
                )
            elif t in ("polyline", "pline"):
                pts = [(float(p[0]), float(p[1])) for p in it["points"]]
                if len(pts) < 2:
                    raise ValueError("polyline needs at least 2 points")
                ms.AddLightWeightPolyline(_coords(pts))
            elif t == "circle":
                ms.AddCircle(
                    _point(float(it["center_x"]), float(it["center_y"])),
                    float(it["radius"]),
                )
            elif t == "arc":
                _add_arc(ms, float(it["center_x"]), float(it["center_y"]),
                         float(it["radius"]), float(it["start_deg"]), float(it["end_deg"]))
            elif t in ("rounded_rectangle", "rounded_rect", "rrect"):
                _add_rounded_rect(ms, float(it["x1"]), float(it["y1"]),
                                  float(it["x2"]), float(it["y2"]), float(it.get("radius", 0)))
            elif t == "text":
                _add_text(ms, it["text"], float(it["x"]), float(it["y"]),
                          float(it.get("height", 2.5)), it.get("align", "left"),
                          float(it.get("rotation", 0.0)))
            elif t == "block":
                scale = float(it.get("scale", 1.0))
                ref = ms.InsertBlock(
                    _point(float(it["x"]), float(it["y"])), str(it["name"]),
                    scale, scale, scale, math.radians(float(it.get("rotation", 0.0))),
                )
                blayer = it.get("layer")
                if blayer:
                    try:
                        ref.Layer = str(blayer)
                    except Exception:
                        pass
                attrs = it.get("attributes")
                if attrs:
                    try:
                        if ref.HasAttributes:
                            for att in ref.GetAttributes():
                                if att.TagString in attrs:
                                    att.TextString = str(attrs[att.TagString])
                    except Exception:
                        pass
            else:
                raise ValueError(f"unknown type '{t}'")
            ok += 1
        except Exception as e:
            errors.append(f"#{i} ({it.get('type', '?')}): {e}")
    msg = f"Drew {ok}/{len(items)} object(s)."
    if errors:
        msg += " Errors: " + "; ".join(errors[:10])
    return msg


@mcp.tool()
def create_layer(name: str, color: int = 7) -> str:
    """Create a layer with the given name and AutoCAD Color Index (1-255; 7 = white)."""
    acad = _get_acad()
    layer = acad.ActiveDocument.Layers.Add(name)
    layer.color = int(color)
    return f"Layer '{name}' created with color index {color}."


@mcp.tool()
def set_active_layer(name: str) -> str:
    """Make the named layer the active layer for new objects."""
    acad = _get_acad()
    doc = acad.ActiveDocument
    doc.ActiveLayer = doc.Layers.Item(name)
    return f"Active layer set to '{name}'."


@mcp.tool()
def activate_layout(name: str = "Model") -> str:
    """Switch the active space/page. Pass "Model" for model space, or a paper-space layout name
    such as "Layout1". This controls what capture_view shows and which page new paper-space work
    goes onto -- useful for multi-page drawing sets. Matching is case-insensitive."""
    acad = _get_acad()
    doc = acad.ActiveDocument
    try:
        layouts = doc.Layouts
        names = [layouts.Item(i).Name for i in range(layouts.Count)]
    except Exception as e:
        return f"Could not read layouts: {e}"
    target = next((n for n in names if n.lower() == name.lower()), None)
    if target is None:
        return f"No layout named '{name}'. Available: {', '.join(names)}"
    try:
        doc.ActiveLayout = layouts.Item(target)
    except Exception as e:
        return f"Could not activate '{target}': {e}"
    return f"Active space is now '{target}'. Available: {', '.join(names)}"


@mcp.tool()
def set_text_style(name: str = "EFF", font: str = "romans.shx", width_factor: float = 0.9) -> str:
    """Create (or update) a text style with the given SHX font and width factor, and make it the
    active style so new text uses it. A width_factor below 1.0 (e.g. 0.85) gives tidier, less
    spread-out lettering. Common fonts: 'romans.shx', 'simplex.shx', 'txt.shx'. Run this once
    before drawing text to get clean, consistent labels."""
    acad = _get_acad()
    doc = acad.ActiveDocument
    styles = doc.TextStyles
    try:
        style = styles.Item(name)
    except Exception:
        style = styles.Add(name)
    try:
        style.fontFile = font
    except Exception as e:
        return f"Style '{name}' created but the font '{font}' could not be set: {e}"
    try:
        style.Width = float(width_factor)
    except Exception:
        pass
    try:
        doc.ActiveTextStyle = style
    except Exception:
        pass
    return f"Active text style set to '{name}' (font {font}, width {width_factor})."


@mcp.tool()
def list_layers() -> str:
    """List all layers in the active drawing."""
    acad = _get_acad()
    layers = acad.ActiveDocument.Layers
    names = [layers.Item(i).Name for i in range(layers.Count)]
    return "Layers: " + ", ".join(names)


@mcp.tool()
def set_layer_lineweight(name: str, lineweight_mm: float) -> str:
    """Set a layer's line thickness (lineweight) in millimetres. Objects drawn on that layer with
    default thickness will display and plot at this weight. Snaps to the nearest standard CAD
    lineweight (e.g. 0.13, 0.25, 0.50, 1.00 mm)."""
    acad = _get_acad()
    layer = acad.ActiveDocument.Layers.Item(name)
    lw = _nearest_lineweight(lineweight_mm)
    layer.Lineweight = lw
    _show_lineweights()
    return f"Layer '{name}' lineweight set to {lw / 100:.2f} mm."


@mcp.tool()
def insert_block(block_name: str, x: float, y: float, scale: float = 1.0,
                 rotation_deg: float = 0.0, layer: str = "",
                 attributes: dict | None = None) -> str:
    """Insert a block reference (symbol) by name at (x, y) with uniform scale and rotation (deg).
    Optionally place it on a given layer and fill in its attribute values, e.g.
    attributes={"TAG": "-K1", "PARTNO": "3RT2015-1AP01"}. Use get_block_attributes(block_name)
    first to see which attribute tags a symbol has."""
    import math
    ms = _model_space()
    ref = ms.InsertBlock(
        _point(x, y), block_name,
        float(scale), float(scale), float(scale),
        math.radians(float(rotation_deg)),
    )
    if layer:
        try:
            ref.Layer = layer
        except Exception:
            pass
    set_count = 0
    if attributes:
        try:
            if ref.HasAttributes:
                for att in ref.GetAttributes():
                    if att.TagString in attributes:
                        att.TextString = str(attributes[att.TagString])
                        set_count += 1
        except Exception:
            pass
    extra = f", set {set_count} attribute(s)" if attributes else ""
    on_layer = f" on layer '{layer}'" if layer else ""
    return f"Block '{block_name}' inserted at ({x}, {y}){on_layer}{extra}."


@mcp.tool()
def list_blocks() -> str:
    """List the named block definitions (reusable symbols) defined in the drawing, so they can be
    placed with insert_block. Skips internal/anonymous blocks (names starting with '*' such as
    model space, paper space, and hatch helpers)."""
    acad = _get_acad()
    blocks = acad.ActiveDocument.Blocks
    names = []
    for i in range(blocks.Count):
        try:
            name = blocks.Item(i).Name
        except Exception:
            continue
        if name.startswith("*"):
            continue
        names.append(name)
    names.sort()
    if not names:
        return "No named blocks (reusable symbols) found in this drawing."
    return f"{len(names)} block(s): " + ", ".join(names)


@mcp.tool()
def get_block_attributes(block_name: str) -> str:
    """List the attribute tags (editable text fields, e.g. a component tag or part number) defined
    on a block, so you know what values insert_block can fill in for that symbol."""
    acad = _get_acad()
    try:
        block = acad.ActiveDocument.Blocks.Item(block_name)
    except Exception:
        return f"No block named '{block_name}' in this drawing."
    tags = []
    for i in range(block.Count):
        try:
            e = block.Item(i)
            if "AttributeDefinition" in e.ObjectName:
                tags.append(e.TagString)
        except Exception:
            pass
    if not tags:
        return f"Block '{block_name}' has no attributes (it's a fixed symbol)."
    return f"Block '{block_name}' attributes: " + ", ".join(tags)


@mcp.tool()
def zoom_extents() -> str:
    """Zoom the AutoCAD view to fit all drawing objects."""
    _get_acad().ZoomExtents()
    return "Zoomed to extents."


@mcp.tool()
def save_drawing(path: str = "") -> str:
    """Save the active drawing. Pass a full .dwg path to Save As, or leave empty to Save."""
    doc = _get_acad().ActiveDocument
    if path:
        doc.SaveAs(path)
        return f"Drawing saved as {path}."
    doc.Save()
    return "Drawing saved."


@mcp.tool()
def list_entities(limit: int = 200) -> str:
    """List the objects in the active drawing's model space with their type and key geometry
    (line endpoints, circle center/radius, polyline vertices, text content). Use this to read
    what already exists in the drawing."""
    import json
    ms = _model_space()
    total = ms.Count
    shown = min(total, max(1, limit))

    def pl(p):
        return [round(float(c), 4) for c in p]

    items = []
    for i in range(shown):
        e = ms.Item(i)
        try:
            kind = e.ObjectName
        except Exception:
            kind = "Unknown"
        info = {"index": i, "type": kind}
        try:
            if "Circle" in kind:
                info["center"] = pl(e.Center)
                info["radius"] = round(float(e.Radius), 4)
            elif "Arc" in kind:
                info["center"] = pl(e.Center)
                info["radius"] = round(float(e.Radius), 4)
            elif "Polyline" in kind:
                coords = list(e.Coordinates)
                info["vertices"] = [
                    [round(float(coords[j]), 4), round(float(coords[j + 1]), 4)]
                    for j in range(0, len(coords) - 1, 2)
                ]
            elif "Line" in kind:
                info["start"] = pl(e.StartPoint)
                info["end"] = pl(e.EndPoint)
            elif "Text" in kind:
                info["text"] = e.TextString
                info["position"] = pl(e.InsertionPoint)
        except Exception as ex:
            info["note"] = f"details unavailable: {ex}"
        items.append(info)

    header = f"{total} object(s) in model space"
    if shown < total:
        header += f" (showing first {shown})"
    return header + "\n" + json.dumps(items, ensure_ascii=False)


@mcp.tool()
def get_drawing_extents() -> str:
    """Return the overall bounding box (min and max corners) and overall width/height of all
    geometry in the active drawing, in drawing units. Useful for the overall size of a part."""
    ms = _model_space()
    minx = miny = float("inf")
    maxx = maxy = float("-inf")
    found = False
    for i in range(ms.Count):
        try:
            lo, hi = ms.Item(i).GetBoundingBox()
            lo = [float(c) for c in lo]
            hi = [float(c) for c in hi]
            minx, miny = min(minx, lo[0]), min(miny, lo[1])
            maxx, maxy = max(maxx, hi[0]), max(maxy, hi[1])
            found = True
        except Exception:
            pass
    if not found:
        return "No measurable objects found in the drawing."
    return (
        f"min=({minx:.3f}, {miny:.3f}), max=({maxx:.3f}, {maxy:.3f}), "
        f"width={maxx - minx:.3f}, height={maxy - miny:.3f} (drawing units)"
    )


@mcp.tool()
def capture_view():
    """Take a screenshot of the live ZWCAD/AutoCAD window and return it as an image, so you can see
    what is currently on screen in the drawing. Zoom/pan in the CAD window first to frame the area
    you want to look at, then call this to inspect it, locate circuits, or check your own work."""
    import io
    import time
    acad = _get_acad()

    try:
        from PIL import ImageGrab
    except Exception:
        return ("Screenshot support isn't installed on this machine. In the repo folder run: "
                "  .venv\\Scripts\\activate ; uv pip install -e .   then fully restart Claude.")

    img = None
    try:
        import win32gui
        import win32con
        from ctypes import windll
        try:
            windll.user32.SetProcessDPIAware()
        except Exception:
            pass
        hwnd = int(acad.HWND)
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        # Force the CAD window above everything (including this app) so the grab captures the
        # drawing, not whatever window happens to be in front of it.
        flags = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, flags)
        time.sleep(0.4)
        rect = win32gui.GetWindowRect(hwnd)
        img = ImageGrab.grab(bbox=rect)
        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, flags)
    except Exception:
        img = None

    if img is None:
        try:
            img = ImageGrab.grab()  # fallback: whole screen
        except Exception as e:
            return f"Could not capture the screen: {e}"

    try:
        from mcp.server.fastmcp import Image
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return Image(data=buf.getvalue(), format="png")
    except Exception as e:
        return f"Could not encode the screenshot: {e}"


@mcp.tool()
def get_selected_entities() -> str:
    """Report the objects currently selected in the CAD window. Select them first (window-select so
    they show grips), then run this. Lists each object's type, layer, and key geometry -- block
    symbols (with name + position + rotation), text, lines, circles, polylines. Use this to capture
    exactly how a specific circuit or region of a drawing is built."""
    import json
    acad = _get_acad()
    doc = acad.ActiveDocument
    try:
        sel = doc.PickfirstSelectionSet
        total = sel.Count
    except Exception as e:
        return f"Could not read the selection: {e}"
    if total == 0:
        return ("Nothing is selected. In the CAD window, window-select the objects you want (so they "
                "show grips/highlight), then run get_selected_entities again.")

    def pl(p):
        return [round(float(c), 2) for c in p]

    items = []
    for i in range(total):
        e = sel.Item(i)
        try:
            kind = e.ObjectName
        except Exception:
            kind = "Unknown"
        info = {"type": kind}
        try:
            info["layer"] = e.Layer
        except Exception:
            pass
        try:
            if "BlockReference" in kind:
                info["block"] = e.Name
                info["position"] = pl(e.InsertionPoint)
                try:
                    info["rotation_deg"] = round(float(e.Rotation) * 180 / 3.141592653589793, 1)
                except Exception:
                    pass
            elif "Circle" in kind:
                info["center"] = pl(e.Center)
                info["radius"] = round(float(e.Radius), 2)
            elif "Polyline" in kind:
                coords = list(e.Coordinates)
                info["vertices"] = [
                    [round(float(coords[j]), 2), round(float(coords[j + 1]), 2)]
                    for j in range(0, len(coords) - 1, 2)
                ]
            elif "Line" in kind:
                info["start"] = pl(e.StartPoint)
                info["end"] = pl(e.EndPoint)
            elif "Text" in kind:
                info["text"] = e.TextString
                info["position"] = pl(e.InsertionPoint)
        except Exception as ex:
            info["note"] = f"details unavailable: {ex}"
        items.append(info)
    return f"{total} object(s) selected:\n" + json.dumps(items, ensure_ascii=False)


@mcp.tool()
def delete_last(count: int = 1) -> str:
    """Delete the most recently added object(s) from the drawing. Handy for undoing shapes that
    were just drawn. `count` is how many of the newest objects to remove."""
    ms = _model_space()
    total = ms.Count
    n = min(max(1, count), total)
    deleted = 0
    for i in range(total - 1, total - 1 - n, -1):
        try:
            ms.Item(i).Delete()
            deleted += 1
        except Exception:
            pass
    return f"Deleted {deleted} of the most recent object(s). {ms.Count} object(s) remain."


@mcp.tool()
def delete_entities(indices: list[int]) -> str:
    """Delete specific objects by their model-space index (the index shown by list_entities),
    e.g. [3, 4, 5]. Re-run list_entities first if the drawing changed, since indices shift
    after a delete."""
    ms = _model_space()
    count = ms.Count
    valid = sorted({i for i in indices if 0 <= i < count}, reverse=True)
    deleted = 0
    for i in valid:
        try:
            ms.Item(i).Delete()
            deleted += 1
        except Exception:
            pass
    return f"Deleted {deleted} of {len(indices)} requested object(s). {ms.Count} object(s) remain."


@mcp.tool()
def hatch_region(boundary_indices: list[int], spacing: float, angle_deg: float = 90.0) -> str:
    """Fill a closed region with evenly spaced parallel lines (e.g. deck planking), automatically
    clipped to the region's outline so the lines start and stop exactly at the boundary.
    `boundary_indices` are the index/indices (from list_entities) of the closed outline: one closed
    polyline, or several segments that together form a closed loop. `spacing` is the gap between
    lines in drawing units; angle_deg=90 gives vertical lines, 0 gives horizontal."""
    import math
    import pythoncom
    import win32com.client
    ms = _model_space()
    try:
        boundaries = [ms.Item(i) for i in boundary_indices]
        loop = win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, boundaries
        )
        hatch = ms.AddHatch(0, "_USER", True)  # 0 = user-defined parallel lines
        hatch.AppendOuterLoop(loop)            # must be the first call after AddHatch
        hatch.PatternSpace = float(spacing)
        hatch.PatternAngle = math.radians(float(angle_deg))
        hatch.Evaluate()
        try:
            _get_acad().ActiveDocument.Regen(True)
        except Exception:
            pass
        return (
            f"Filled the region (boundary {boundary_indices}) with lines {spacing} units apart "
            f"at {angle_deg} deg, clipped to the outline."
        )
    except Exception as e:
        return (
            f"Could not create the hatch: {e}. The boundary must be a single closed shape, or "
            f"connected segments that form one closed loop. Use list_entities to find the right "
            f"index, and make sure the outline is actually closed."
        )


@mcp.tool()
def draw_arc(center_x: float, center_y: float, radius: float,
             start_deg: float, end_deg: float) -> str:
    """Draw an arc centred at (center_x, center_y), sweeping counter-clockwise from start_deg to
    end_deg. Angles are in degrees measured from the +X axis (0 = east, 90 = north). Use for
    rounded corners, curved symbol parts, cable bends, etc."""
    ms = _model_space()
    _add_arc(ms, center_x, center_y, radius, start_deg, end_deg)
    return f"Arc drawn at ({center_x}, {center_y}) r={radius}, {start_deg} to {end_deg} deg."


@mcp.tool()
def draw_rounded_rectangle(x1: float, y1: float, x2: float, y2: float, radius: float) -> str:
    """Draw a rectangle with rounded corners between corners (x1, y1) and (x2, y2). `radius` is the
    corner fillet radius (e.g. 30 for the R30 remote-control cut-out). The radius is clamped so it
    never exceeds half the shorter side. Ideal for panel outlines, cut-outs and device bezels."""
    ms = _model_space()
    _add_rounded_rect(ms, x1, y1, x2, y2, radius)
    return f"Rounded rectangle drawn from ({x1}, {y1}) to ({x2}, {y2}), corner radius {radius}."


@mcp.tool()
def add_dimension(x1: float, y1: float, x2: float, y2: float,
                  dim_line_x: float, dim_line_y: float, direction: str = "aligned") -> str:
    """Add a dimension measuring between (x1, y1) and (x2, y2). (dim_line_x, dim_line_y) is a point
    the dimension line passes through -- offset it away from the part so the dimension sits clear.
    `direction`: "horizontal" measures the X distance, "vertical" the Y distance, "aligned" the
    true straight-line distance (parallel to the two points). Use this to reproduce dimensioned
    sheets such as the mounting drawing (600 x 1000 cabinet, the cut-out, etc.)."""
    import math
    ms = _model_space()
    p1, p2, dl = _point(x1, y1), _point(x2, y2), _point(dim_line_x, dim_line_y)
    d = str(direction).lower()
    try:
        if d.startswith("h"):
            ms.AddDimRotated(p1, p2, dl, 0.0)
        elif d.startswith("v"):
            ms.AddDimRotated(p1, p2, dl, math.radians(90.0))
        else:
            ms.AddDimAligned(p1, p2, dl)
    except Exception as e:
        return f"Could not create the dimension: {e}"
    return f"{direction} dimension added between ({x1}, {y1}) and ({x2}, {y2})."


@mcp.tool()
def add_mtext(text: str, x: float, y: float, width: float = 100.0, height: float = 2.5) -> str:
    """Place a multi-line (paragraph) text box with its top-left corner at (x, y). `width` is the
    wrapping width in drawing units; `height` is the character height. Use \\n in `text` for line
    breaks. Good for title-block notes and longer labels that a single-line add_text can't hold."""
    ms = _model_space()
    try:
        mt = ms.AddMText(_point(x, y), float(width), str(text).replace("\\n", "\n").replace("\n", "\\P"))
        try:
            mt.Height = float(height)
        except Exception:
            pass
    except Exception as e:
        return f"Could not create the mtext: {e}"
    return f"MText placed at ({x}, {y}), width {width}, height {height}."


@mcp.tool()
def define_block(name: str, indices: list[int], base_x: float, base_y: float,
                 erase_source: bool = True) -> str:
    """Turn existing model-space objects into a reusable block (symbol) named `name`, so it can be
    stamped repeatedly with insert_block / a "block" item in draw_batch. `indices` are the
    model-space indices (from list_entities) of the objects to capture; (base_x, base_y) is the
    block's insertion/origin point. By default the originals are erased after the block is made
    (set erase_source=false to keep them). This is the key to schematic sheets: draw one relay /
    terminal / monitor symbol, capture it once, then place it everywhere with consistent geometry."""
    import pythoncom
    import win32com.client
    acad = _get_acad()
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    count = ms.Count
    sources = []
    for i in indices:
        if 0 <= i < count:
            try:
                sources.append(ms.Item(i))
            except Exception:
                pass
    if not sources:
        return f"No valid objects at indices {indices} (drawing has {count} objects)."
    try:
        blk = doc.Blocks.Add(_point(base_x, base_y), name)
    except Exception as e:
        return f"Could not create block '{name}': {e}"
    try:
        arr = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, sources)
        doc.CopyObjects(arr, blk)
    except Exception as e:
        return (f"Block '{name}' was created but copying the {len(sources)} object(s) into it "
                f"failed: {e}")
    erased = 0
    if erase_source:
        for e in sources:
            try:
                e.Delete()
                erased += 1
            except Exception:
                pass
    tail = f", erased {erased} source object(s)" if erase_source else ""
    return (f"Block '{name}' defined from {len(sources)} object(s) at base "
            f"({base_x}, {base_y}){tail}. Place it with insert_block('{name}', x, y).")


def _add_dot(ms, x, y, radius=1.0):
    """Add a filled connection dot (a small circle flooded with a SOLID hatch)."""
    import pythoncom
    import win32com.client
    c = ms.AddCircle(_point(x, y), float(radius))
    try:
        loop = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [c])
        h = ms.AddHatch(1, "SOLID", True)  # 1 = predefined pattern
        h.AppendOuterLoop(loop)
        h.Evaluate()
    except Exception:
        pass  # leave the outline if the fill fails
    return c


@mcp.tool()
def get_block_geometry(name: str, limit: int = 400) -> str:
    """Read the geometry *inside* a block definition (symbol), in the block's own coordinate system.
    This is how you find a symbol's connection points before wiring to it: the library symbols carry
    no attribute pins, so their terminals are the free ends of the internal lines. Returns each
    primitive with coordinates (line start/end, circle/arc centre+radius, polyline vertices, text /
    attribute-definition positions) plus the block's base point. Place the block with insert_block at
    (X, Y); a point (px, py) listed here then lands near (X + px, Y + py) at scale 1, rotation 0."""
    import json
    acad = _get_acad()
    doc = acad.ActiveDocument
    try:
        blk = doc.Blocks.Item(name)
    except Exception as e:
        return f"No block named '{name}': {e}. Use list_blocks to see the available symbols."

    def pl(p):
        return [round(float(c), 4) for c in p]

    base = None
    try:
        base = pl(blk.Origin)
    except Exception:
        pass
    total = blk.Count
    shown = min(total, max(1, int(limit)))
    deg = 180.0 / 3.141592653589793
    items = []
    for i in range(shown):
        e = blk.Item(i)
        try:
            kind = e.ObjectName
        except Exception:
            kind = "Unknown"
        info = {"type": kind}
        try:
            if "Arc" in kind:
                info["center"] = pl(e.Center)
                info["radius"] = round(float(e.Radius), 4)
                info["start_deg"] = round(float(e.StartAngle) * deg, 2)
                info["end_deg"] = round(float(e.EndAngle) * deg, 2)
            elif "Circle" in kind:
                info["center"] = pl(e.Center)
                info["radius"] = round(float(e.Radius), 4)
            elif "Polyline" in kind:
                c = list(e.Coordinates)
                info["vertices"] = [
                    [round(float(c[j]), 4), round(float(c[j + 1]), 4)]
                    for j in range(0, len(c) - 1, 2)
                ]
            elif "Line" in kind:
                info["start"] = pl(e.StartPoint)
                info["end"] = pl(e.EndPoint)
            elif "AttributeDefinition" in kind:
                info["tag"] = e.TagString
                info["position"] = pl(e.InsertionPoint)
            elif "Text" in kind:
                info["text"] = e.TextString
                info["position"] = pl(e.InsertionPoint)
            elif "BlockReference" in kind:
                info["block"] = e.Name
                info["position"] = pl(e.InsertionPoint)
            elif "Point" in kind:
                info["position"] = pl(e.Coordinates)
        except Exception as ex:
            info["note"] = f"details unavailable: {ex}"
        items.append(info)

    header = f"Block '{name}': {total} entit(ies)"
    if base is not None:
        header += f", base point {base[:2]}"
    if shown < total:
        header += f" (showing first {shown})"
    header += ". Free line ends and attribute positions are usually the terminals."
    return header + "\n" + json.dumps(items, ensure_ascii=False)


@mcp.tool()
def get_entity_bounds(indices: list[int]) -> str:
    """Return the bounding box (min/max corner + width/height) of model-space objects by index
    (from list_entities), e.g. [12, 13]. Works for block references too -- use it to find where a
    placed symbol actually sits and how much room it takes, so you can route wires clear of it."""
    import json
    ms = _model_space()
    count = ms.Count
    out = []
    for i in indices:
        if not (0 <= i < count):
            out.append({"index": i, "note": "out of range"})
            continue
        try:
            lo, hi = ms.Item(i).GetBoundingBox()
            lo = [round(float(c), 3) for c in lo]
            hi = [round(float(c), 3) for c in hi]
            out.append({
                "index": i, "min": lo[:2], "max": hi[:2],
                "width": round(hi[0] - lo[0], 3), "height": round(hi[1] - lo[1], 3),
            })
        except Exception as e:
            out.append({"index": i, "note": f"no bounds: {e}"})
    return json.dumps(out, ensure_ascii=False)


@mcp.tool()
def draw_wire(x1: float, y1: float, x2: float, y2: float,
              route: str = "auto", dots: bool = False, dot_radius: float = 1.0) -> str:
    """Draw a schematic wire from (x1, y1) to (x2, y2) as a polyline. `route`: "hv" goes horizontal
    then vertical (corner at x2,y1), "vh" goes vertical then horizontal (corner at x1,y2), "direct"
    is a straight segment, "auto" picks an L-bend when the points differ in both axes (else straight).
    Set dots=True to drop a filled junction dot at each end (radius `dot_radius`)."""
    ms = _model_space()
    x1, y1, x2, y2 = float(x1), float(y1), float(x2), float(y2)
    r = str(route).lower()
    if r == "direct" or x1 == x2 or y1 == y2:
        pts = [(x1, y1), (x2, y2)]
    elif r == "vh":
        pts = [(x1, y1), (x1, y2), (x2, y2)]
    else:  # "hv" or "auto"
        pts = [(x1, y1), (x2, y1), (x2, y2)]
    ms.AddLightWeightPolyline(_coords(pts))
    if dots:
        _add_dot(ms, x1, y1, dot_radius)
        _add_dot(ms, x2, y2, dot_radius)
    return f"Wire drawn ({route}) from ({x1}, {y1}) to ({x2}, {y2})."


@mcp.tool()
def draw_dot(x: float, y: float, radius: float = 1.0) -> str:
    """Place a filled connection/junction dot at (x, y). Use at points where wires join (a tee),
    so the connection reads as intentional rather than a crossing."""
    ms = _model_space()
    _add_dot(ms, x, y, radius)
    return f"Connection dot placed at ({x}, {y}), radius {radius}."


@mcp.tool()
def copy_entities(indices: list[int], dx: float, dy: float) -> str:
    """Copy model-space objects (by index from list_entities) by an offset (dx, dy). Good for
    repeating a rung, a terminal, or a whole sub-circuit across a sheet. Originals are kept; the
    copies are added at the end of the drawing."""
    ms = _model_space()
    count = ms.Count
    made = 0
    for i in indices:
        if 0 <= i < count:
            try:
                cp = ms.Item(i).Copy()
                cp.Move(_point(0, 0, 0), _point(float(dx), float(dy), 0))
                made += 1
            except Exception:
                pass
    return f"Copied {made} of {len(indices)} object(s) by ({dx}, {dy}). {ms.Count} object(s) now."


@mcp.tool()
def mirror_entities(indices: list[int], x1: float, y1: float, x2: float, y2: float,
                    keep_original: bool = True) -> str:
    """Mirror model-space objects (by index) about the line through (x1, y1)-(x2, y2). Ideal for the
    reversing contactor (draw KM1, mirror it to KM1's twin) or a symmetric panel half. By default the
    original is kept; set keep_original=false to flip in place."""
    ms = _model_space()
    count = ms.Count
    p1, p2 = _point(float(x1), float(y1)), _point(float(x2), float(y2))
    made = 0
    originals = []
    for i in indices:
        if 0 <= i < count:
            try:
                e = ms.Item(i)
                e.Mirror(p1, p2)
                made += 1
                if not keep_original:
                    originals.append(e)
            except Exception:
                pass
    for e in originals:
        try:
            e.Delete()
        except Exception:
            pass
    return f"Mirrored {made} of {len(indices)} object(s) about ({x1},{y1})-({x2},{y2})."


# ===========================================================================
# BUILT-IN SYMBOL LIBRARY
# ---------------------------------------------------------------------------
# A self-contained set of IEC electrical symbols, defined as block definitions
# directly via COM (doc.Blocks.Add + drawing into the block). This means you can
# draw a full schematic in ANY drawing -- including a brand-new blank file with
# no blocks in it -- by calling load_symbol_library() once, then placing the
# EFF_* symbols with insert_block / draw_batch.
#
# These are generic IEC symbols (not the MARSIS house blocks). All multi-pole
# power devices use an 8.6-unit pole pitch so they align on the same 3-phase bus.
# Terminal coordinates (relative to the insertion point) are in _SYMBOL_TERMINALS
# and reported by symbol_info().
# ===========================================================================

POLE = 8.6  # pole-to-pole pitch for 3-phase power symbols (L1, L2, L3 columns)


# -- tiny helpers that draw INTO a block-definition container ---------------
# A Block object exposes the same Add* methods as ModelSpace, so we reuse the
# server's _point / _coords / _add_text helpers and pass the block as container.
def _bl(blk, x1, y1, x2, y2):
    blk.AddLine(_point(x1, y1), _point(x2, y2))


def _bp(blk, pts, closed=False):
    poly = blk.AddLightWeightPolyline(_coords([(p[0], p[1]) for p in pts]))
    if closed:
        poly.Closed = True
    return poly


def _bc(blk, cx, cy, r):
    blk.AddCircle(_point(cx, cy), float(r))


def _bt(blk, text, x, y, height, align="center"):
    _add_text(blk, text, x, y, height, align)


# -- symbol builders --------------------------------------------------------
# Origin (0,0) is the block insertion point. Power symbols: L1 at x=0,
# L2 at x=8.6, L3 at x=17.2; line side at the top (y=0), load side at the bottom.

def _contactor_3p(blk):
    """3-pole contactor main (NO) contacts. Top 1/3/5 at y=0, bottom 2/4/6 at y=-30."""
    for px in (0.0, POLE, 2 * POLE):
        _bl(blk, px, 0, px, -10)          # top terminal stub
        _bl(blk, px, -20, px, -30)        # bottom terminal stub
        _bl(blk, px, -20, px - 4, -10)    # NO moving contact (open, angled)
    _bl(blk, -4, -10, 2 * POLE - 4, -10)  # mechanical link across the blades


def _mcb(blk, poles):
    """Miniature circuit breaker, `poles` poles. Top at y=0, bottom at y=-30."""
    span = (poles - 1) * POLE
    for k in range(poles):
        px = k * POLE
        _bl(blk, px, 0, px, -7)                                   # top stub
        _bl(blk, px, -7, px + 4, -13)                             # switch blade
        _bp(blk, [(px - 2, -15), (px + 2, -15),
                  (px + 2, -22), (px - 2, -22)], closed=True)      # thermal element
        _bl(blk, px, -22, px, -30)                                # bottom stub
    if poles > 1:
        _bl(blk, 0, -10, span, -10)                               # ganged link


def _motor_prot_3p(blk):
    """3RV-style motor protective breaker, 3 poles. Top y=0, bottom y=-35."""
    for px in (0.0, POLE, 2 * POLE):
        _bl(blk, px, 0, px, -8)
        _bp(blk, [(px - 2.5, -8), (px + 2.5, -8),
                  (px + 2.5, -18), (px - 2.5, -18)], closed=True)  # thermal-mag body
        _bl(blk, px - 2.5, -13, px + 2.5, -13)                     # element divider
        _bl(blk, px, -18, px, -35)
    _bl(blk, 0, -11, 2 * POLE, -11)                                # ganged link


def _motor_3ph(blk):
    """3-phase motor. Terminals U/V/W at (0,0), (8.6,0), (17.2,0); circle below."""
    for px in (0.0, POLE, 2 * POLE):
        _bl(blk, px, 0, px, -6)
        _bl(blk, px, -6, POLE, -10)          # converge to circle top
    _bc(blk, POLE, -22, 12)
    _bt(blk, "M", POLE, -19, 6)
    _bt(blk, "3~", POLE, -27, 3.5)


def _phase_monitor(blk):
    """Phase-sequence/loss relay. 3 phase terminals on top, 2 output on bottom."""
    _bp(blk, [(-3, 3), (2 * POLE + 3, 3),
              (2 * POLE + 3, -22), (-3, -22)], closed=True)        # body box
    for px in (0.0, POLE, 2 * POLE):
        _bl(blk, px, 3, px, 7)                                     # phase terminal stubs
    _bl(blk, POLE - 4, -22, POLE - 4, -26)                         # output stub 1
    _bl(blk, POLE + 4, -22, POLE + 4, -26)                         # output stub 2
    _bt(blk, "U<", POLE, -6, 5)
    _bt(blk, "3~", POLE, -14, 3)


def _coil(blk):
    """Contactor/relay coil (A1 top, A2 bottom)."""
    _bp(blk, [(0, 0), (8, 0), (8, -6), (0, -6)], closed=True)
    _bl(blk, 4, 0, 4, 4)        # A1 stub up
    _bl(blk, 4, -6, 4, -10)     # A2 stub down


def _contact_no(blk):
    """NO auxiliary/control contact. Terminals (0,0) and (0,-12)."""
    _bl(blk, 0, 0, 0, -3)       # top fixed
    _bl(blk, 0, -9, 0, -12)     # bottom fixed
    _bl(blk, 0, -9, -4, -2)     # moving blade (open)


def _contact_nc(blk):
    """NC auxiliary/control contact. Terminals (0,0) and (0,-12)."""
    _bl(blk, 0, 0, 0, -3)
    _bl(blk, 0, -9, 0, -12)
    _bl(blk, 0, -9, -4, -2)     # moving blade
    _bl(blk, -4, -2, -4, -5)    # NC bar


def _limit_switch(blk):
    """Travel/torque limit switch: a NO contact with an actuating lever."""
    _contact_no(blk)
    _bl(blk, -4, -2, -8, 1)     # lever


def _terminal(blk):
    """Single field terminal (klem)."""
    _bc(blk, 0, 0, 1.3)


# -- registry + terminal map ------------------------------------------------
_SYMBOLS = {
    "EFF_CONTACTOR_3P":  _contactor_3p,
    "EFF_MCB_1P":        lambda b: _mcb(b, 1),
    "EFF_MCB_2P":        lambda b: _mcb(b, 2),
    "EFF_MCB_3P":        lambda b: _mcb(b, 3),
    "EFF_MOTORPROT_3P":  _motor_prot_3p,
    "EFF_MOTOR_3PH":     _motor_3ph,
    "EFF_PHASE_MONITOR": _phase_monitor,
    "EFF_COIL":          _coil,
    "EFF_CONTACT_NO":    _contact_no,
    "EFF_CONTACT_NC":    _contact_nc,
    "EFF_LIMIT_SWITCH":  _limit_switch,
    "EFF_TERMINAL":      _terminal,
}

# Terminal connection points relative to the insertion point (for wiring).
_SYMBOL_TERMINALS = {
    "EFF_CONTACTOR_3P":  {"top(1,3,5)": [(0, 0), (POLE, 0), (2 * POLE, 0)],
                          "bottom(2,4,6)": [(0, -30), (POLE, -30), (2 * POLE, -30)]},
    "EFF_MCB_3P":        {"top(1,3,5)": [(0, 0), (POLE, 0), (2 * POLE, 0)],
                          "bottom(2,4,6)": [(0, -30), (POLE, -30), (2 * POLE, -30)]},
    "EFF_MCB_2P":        {"top": [(0, 0), (POLE, 0)], "bottom": [(0, -30), (POLE, -30)]},
    "EFF_MCB_1P":        {"top": [(0, 0)], "bottom": [(0, -30)]},
    "EFF_MOTORPROT_3P":  {"top": [(0, 0), (POLE, 0), (2 * POLE, 0)],
                          "bottom": [(0, -35), (POLE, -35), (2 * POLE, -35)]},
    "EFF_MOTOR_3PH":     {"U/V/W": [(0, 0), (POLE, 0), (2 * POLE, 0)]},
    "EFF_PHASE_MONITOR": {"phases": [(0, 7), (POLE, 7), (2 * POLE, 7)],
                          "output": [(POLE - 4, -26), (POLE + 4, -26)]},
    "EFF_COIL":          {"A1": [(4, 4)], "A2": [(4, -10)]},
    "EFF_CONTACT_NO":    {"terminals": [(0, 0), (0, -12)]},
    "EFF_CONTACT_NC":    {"terminals": [(0, 0), (0, -12)]},
    "EFF_LIMIT_SWITCH":  {"terminals": [(0, 0), (0, -12)]},
    "EFF_TERMINAL":      {"point": [(0, 0)]},
}


def _block_exists(doc, name):
    try:
        doc.Blocks.Item(name)
        return True
    except Exception:
        return False


@mcp.tool()
def load_symbol_library(overwrite: bool = False) -> str:
    """Create the built-in IEC electrical symbol library as block definitions in the active
    drawing, so they can be placed with insert_block / draw_batch even in a blank file. Symbols:
    EFF_CONTACTOR_3P, EFF_MCB_1P/2P/3P, EFF_MOTORPROT_3P, EFF_MOTOR_3PH, EFF_PHASE_MONITOR,
    EFF_COIL, EFF_CONTACT_NO, EFF_CONTACT_NC, EFF_LIMIT_SWITCH, EFF_TERMINAL. Run once per drawing
    (or save the drawing as a .dwt template so new drawings already include them). Set overwrite=True
    to rebuild symbols that already exist. Use symbol_info() to see terminal coordinates for wiring."""
    acad = _get_acad()
    doc = acad.ActiveDocument

    # Build symbol geometry on layer "0" so each placed block inherits the
    # insertion layer's colour/lineweight (standard block practice).
    prev_layer = None
    try:
        prev_layer = doc.ActiveLayer
        doc.ActiveLayer = doc.Layers.Item("0")
    except Exception:
        pass

    created, skipped, errors = [], [], []
    for name, builder in _SYMBOLS.items():
        try:
            if _block_exists(doc, name):
                if not overwrite:
                    skipped.append(name)
                    continue
                try:
                    doc.Blocks.Item(name).Delete()  # fails if already referenced
                except Exception:
                    skipped.append(name + "(in use)")
                    continue
            blk = doc.Blocks.Add(_point(0, 0), name)
            builder(blk)
            created.append(name)
        except Exception as e:
            errors.append(f"{name}: {e}")

    if prev_layer is not None:
        try:
            doc.ActiveLayer = prev_layer
        except Exception:
            pass

    msg = f"Symbol library: created {len(created)}"
    if created:
        msg += " (" + ", ".join(created) + ")"
    if skipped:
        msg += f"; skipped {len(skipped)} existing"
    if errors:
        msg += "; errors: " + "; ".join(errors[:8])
    return msg


@mcp.tool()
def symbol_info(name: str = "") -> str:
    """List the built-in library symbols, or -- if `name` is given -- the terminal connection points
    of one symbol (relative to its insertion point, in drawing units), so you know exactly where to
    attach wires after placing it."""
    import json
    if not name:
        return "Library symbols: " + ", ".join(_SYMBOLS.keys())
    if name not in _SYMBOL_TERMINALS:
        return f"No symbol '{name}'. Available: " + ", ".join(_SYMBOLS.keys())
    return f"{name} terminals (relative to insertion point):\n" + json.dumps(
        _SYMBOL_TERMINALS[name], ensure_ascii=False
    )


# ===========================================================================
# STANDARD MACRO  -  Bernard AQ valve control (power + PLC/HMI), one call
# ---------------------------------------------------------------------------
# Wired to the Bernard AQ SWITCH datasheet interface (Start-Up Guide SUG_17003,
# Ch.11.1 "wiring diagram without positioner option"):
#   E  = three-phase motor          S1 = counter-clockwise (normally OPEN)
#   B  = travel limit switches      S2 = clockwise (normally CLOSE)
#   C  = signalling switches        A  = torque limit switches
#   D  = heater resistance (230VAC) + motor thermal protector (thermostat, NC)
# Per SUG_17003 7.3 the travel-limit, torque and thermostat contacts MUST be
# integrated into the control system: they are wired into the contactor coil
# rungs (hardware stop), while the signalling switches feed the PLC for position.
# The SWITCH type has NO automatic phase correction, so reversing = swap two
# phases and a phase-sequence monitor is used as a permissive.
# NOTE: the terminal NUMBERS below are now baked in from Bernard datasheet
# TEC01-03 rev06B, sheet 3.2 (AQ SWITCH, 3-phase). The only thing still marked
# "TBC" is the NO/NC polarity WITHIN each 3-wire switch group, which is set on
# the order-specific sheet inside the unit cover.
# ===========================================================================


# ===========================================================================
# DATASHEET TABLES  -  Bernard AQ range, TEC01-03_E+F_GRP_rev06B
# Baked in so the schematic carries real wiring-sheet terminal numbers and so a
# client can look mechanical/electrical figures up over MCP (aq_model_data).
# Locked source for wiring = sheet 3.2 (SWITCH, 3-phase, no positioner).
# ===========================================================================

# --- Terminal map: sheet 3.2 "AQ SWITCH: 3-phases / Triphase" ---------------
# Each travel/torque switch is a 3-wire group: common + two contacts. Which of
# the pair is NO vs NC ("polarity") is order-specific -> labelled "TBC".
AQ_SWITCH_TERMINALS = {
    "source": "Bernard TEC01-03 rev06B, sheet 3.2 (SWITCH 3-phase); identical layout reprinted on sheet 3.21 for AQ150-1000 SWITCH, so this map covers the whole SWITCH range",
    "motor_3ph":        {"terminals": [1, 2, 3], "pe": "PE",
                         "note": "3Ph direct wiring = Closing (swap 2 phases to reverse)"},
    "motor_thermostat": {"terminals": [40, 41], "type": "NC",
                         "note": "Th* motor thermal protection - wire as a STOP"},
    "torque_open":      {"terminals": [4, 5, 6],    "polarity": "TBC",
                         "label": "Limiteur d'effort Ouvert / Torque limit OPEN"},
    "torque_close":     {"terminals": [7, 8, 9],    "polarity": "TBC",
                         "label": "Limiteur d'effort Ferme / Torque limit CLOSE"},
    "travel_open":      {"terminals": [10, 11, 12], "polarity": "TBC",
                         "label": "Fin de course Ouvert / Travel limit OPEN"},
    "travel_close":     {"terminals": [13, 14, 15], "polarity": "TBC",
                         "label": "Fin de course Ferme / Travel limit CLOSE"},
    "aux_travel_open":  {"terminals": [20, 21, 22], "polarity": "TBC",
                         "label": "Fin de course aux. Ouvert / Extra travel OPEN"},
    "aux_travel_close": {"terminals": [23, 24, 25], "polarity": "TBC",
                         "label": "Fin de course aux. Ferme / Extra travel CLOSE"},
    "heater":           {"terminals": [26, 27],
                         "label": "Resistance de chauffage / Anti-condensation heater"},
    "potentiometer":    {"terminals": [16, 17, 18], "option": True},
    "position_xmitter_4_20mA": {"terminals": {"+": 80, "-": 81},
                                "supply": "12-32 VDC", "option": True},
}

# Single-phase SWITCH (sheet 3.1) kept for completeness.
AQ_SWITCH_TERMINALS_1PH = {
    "source": "Bernard TEC01-03 rev06B, sheet 3.1 (SWITCH 1-phase)",
    "supply":         {"L": 1, "N": "N", "pe": "PE", "open_cmd": 2, "close_cmd": 3},
    "HO_open_travel": 5, "HF_close_travel": 6, "heater": 4,
    "HLF_close_torque": 8, "HLO_open_torque": 7,
    "aux_travel": [20, 21, 22, 23, 24, 25],
    "potentiometer": [16, 17, 18],
    "position_xmitter_4_20mA": {"+": 80, "-": 81, "supply": "12-32 VDC"},
}

# ── Mechanical data: section 2 (all models) ─────────────────────────────────
# switch_kg = SWITCH variant bare; logic_kg = LOGIC variant (integrated controller ~+5 kg)
# flanges = standard ISO5211 flanges; optional_flanges = non-standard option
AQ_DIMENSIONS = {
    "AQ5":   {"switch_kg": 10, "logic_kg": 15, "square_mm": 19, "max_bore_mm": 22,
              "flanges": ["F05","F07"],
              "bolts": {"F05":"4xM6 depth 8mm register d50",
                        "F07":"4xM8 depth 12mm register d70"}},
    "AQ10":  {"switch_kg": 10, "logic_kg": 15, "square_mm": 19, "max_bore_mm": 22,
              "flanges": ["F05","F07"],
              "bolts": {"F05":"4xM6 depth 8mm register d50",
                        "F07":"4xM8 depth 12mm register d70"}},
    "AQ15":  {"switch_kg": 10, "logic_kg": 15, "square_mm": 19, "max_bore_mm": 22,
              "flanges": ["F05","F07"],
              "bolts": {"F05":"4xM6 depth 8mm register d50",
                        "F07":"4xM8 depth 12mm register d70"}},
    "AQ25":  {"switch_kg": 13, "logic_kg": 18, "square_mm": 27, "max_bore_mm": 32,
              "flanges": ["F07","F10"],
              "bolts": {"F07":"4xM8 depth 12mm register d70",
                        "F10":"4xM10 depth 18mm register d102"}},
    "AQ30":  {"switch_kg": 15, "logic_kg": 20, "square_mm": 27, "max_bore_mm": 32,
              "flanges": ["F07","F10"],
              "bolts": {"F07":"4xM8 depth 12mm register d70",
                        "F10":"4xM10 depth 18mm register d102"}},
    "AQ50":  {"switch_kg": 15, "logic_kg": 20, "square_mm": 27, "max_bore_mm": 32,
              "flanges": ["F07","F10"],
              "bolts": {"F07":"4xM8 depth 12mm register d70",
                        "F10":"4xM10 depth 18mm register d102"}},
    "AQ80":  {"switch_kg": 18, "logic_kg": 23, "square_mm": 27, "max_bore_mm": 36,
              "flanges": ["F10","F12"],
              "bolts": {"F10":"4xM10 depth 18mm register d102",
                        "F12":"4xM12 depth 25mm register d125"}},
    "AQ150": {"switch_kg": 38, "logic_kg": 43, "square_mm": 36, "max_bore_mm": 46,
              "flanges": ["F12","F14"], "optional_flanges": ["F10"],
              "bolts": {"F10":"4xM10 depth 18mm register d102",
                        "F12":"4xM12 depth 25mm register d125",
                        "F14":"4xM16 depth 28mm register d140"}},
    "AQ280": {"switch_kg": 50, "logic_kg": 55, "square_mm": 46, "max_bore_mm": 60,
              "flanges": ["F14","F16"], "optional_flanges": ["F12"],
              "bolts": {"F12":"4xM12 depth 25mm register d125",
                        "F14":"4xM16 depth 28mm register d140",
                        "F16":"4xM20 depth 35mm register d160"}},
    "AQ430": {"switch_kg": 79, "logic_kg": 83, "square_mm": 55, "max_bore_mm": 75,
              "flanges": ["F16"],
              "bolts": {"F16":"4xM20 depth 35mm register d160"}},
    "AQ610": {"switch_kg": 86, "logic_kg": 90, "square_mm": 70, "max_bore_mm": 80,
              "flanges": ["F25"],
              "bolts": {"F25":"4xM30 depth 55mm register d225"}},
    "AQ830": {"switch_kg": 94, "logic_kg": 99, "square_mm": 70, "max_bore_mm": 95,
              "flanges": ["F25"], "logic_flanges": ["F16","F20","F25"],
              "bolts": {"F16":"4xM20 depth 35mm register d160",
                        "F20":"4xM24 depth 45mm register d200",
                        "F25":"4xM30 depth 55mm register d225"}},
    "AQ1000":{"switch_kg": 115,"logic_kg": 119,"square_mm": 70, "max_bore_mm": 95,
              "flanges": ["F25"], "logic_flanges": ["F16","F20","F25"],
              "bolts": {"F16":"4xM20 depth 35mm register d160",
                        "F20":"4xM24 depth 45mm register d200",
                        "F25":"4xM30 depth 55mm register d225"}},
}

# ── Performance tables: section 1, all supply voltages ───────────────────────
# Source: Bernard TEC01-03_E+F_GRP_rev06B pages 4-10
# Values are indicative (no-load). Power = at 33% of max torque.
# Each model entry is a list of speed options [{torque_Nm, time_s, kW, nom_A, start_A}].
# Note: former dict AQ_PERFORMANCE_3x400V_50HZ had mislabelled 3x415V data — corrected here.
AQ_PERFORMANCE = {
    "1x110-115VAC_50Hz": {
        "AQ5":  [{"torque_Nm":  50, "time_s": 16, "kW":0.02,"nom_A":1.1, "start_A":1.4}],
        "AQ10": [{"torque_Nm": 100, "time_s": 25, "kW":0.02,"nom_A":1.1, "start_A":1.4}],
        "AQ15": [{"torque_Nm": 150, "time_s": 30, "kW":0.02,"nom_A":1.3, "start_A":1.6}],
        "AQ25": [{"torque_Nm": 250, "time_s": 30, "kW":0.04,"nom_A":1.5, "start_A":1.7}],
        "AQ30": [{"torque_Nm": 300, "time_s": 35, "kW":0.04,"nom_A":1.5, "start_A":1.7}],
        "AQ50": [{"torque_Nm": 500, "time_s": 35, "kW":0.06,"nom_A":2.5, "start_A":4.2},
                 {"torque_Nm": 500, "time_s": 55, "kW":0.04,"nom_A":1.5, "start_A":1.7}],
        "AQ80": [{"torque_Nm": 800, "time_s": 55, "kW":0.06,"nom_A":2.5, "start_A":4.2}],
        "AQ150":[{"torque_Nm":1140, "time_s": 20, "kW":0.7, "nom_A":8.3, "start_A":33},
                 {"torque_Nm":1500, "time_s": 40, "kW":0.7, "nom_A":8.3, "start_A":33},
                 {"torque_Nm":1500, "time_s":100, "kW":0.3, "nom_A":4.5, "start_A":11}],
        "AQ280":[{"torque_Nm":2800, "time_s": 70, "kW":0.7, "nom_A":8.3, "start_A":33},
                 {"torque_Nm":2800, "time_s":100, "kW":0.7, "nom_A":8.3, "start_A":33},
                 {"torque_Nm":2800, "time_s":140, "kW":0.7, "nom_A":8.3, "start_A":33}],
        "AQ430":[{"torque_Nm":4300, "time_s": 40, "kW":1.8, "nom_A":21,  "start_A":77},
                 {"torque_Nm":4300, "time_s": 70, "kW":1.8, "nom_A":21,  "start_A":77},
                 {"torque_Nm":4300, "time_s":120, "kW":0.5, "nom_A":8.5, "start_A":19}],
        "AQ610":[{"torque_Nm":6100, "time_s": 50, "kW":3.0, "nom_A":39,  "start_A":142},
                 {"torque_Nm":6100, "time_s":100, "kW":1.8, "nom_A":21,  "start_A":77},
                 {"torque_Nm":6100, "time_s":140, "kW":1.8, "nom_A":21,  "start_A":77}],
        "AQ830":[{"torque_Nm":6015, "time_s":115, "kW":0.7, "nom_A":8.3, "start_A":33},
                 {"torque_Nm":8300, "time_s":230, "kW":0.7, "nom_A":8.3, "start_A":33}],
        "AQ1000":[{"torque_Nm":10400,"time_s":190,"kW":0.7, "nom_A":8.3, "start_A":33},
                  {"torque_Nm":10400,"time_s":135,"kW":1.2, "nom_A":15,  "start_A":61},
                  {"torque_Nm":10400,"time_s":100,"kW":1.2, "nom_A":15,  "start_A":61},
                  {"torque_Nm":10400,"time_s": 56,"kW":1.8, "nom_A":21,  "start_A":77}],
    },
    "1x110-115VAC_60Hz": {
        "AQ5":  [{"torque_Nm":  50, "time_s": 13, "kW":0.02,"nom_A":1.1, "start_A":1.4}],
        "AQ10": [{"torque_Nm": 100, "time_s": 21, "kW":0.02,"nom_A":1.1, "start_A":1.4}],
        "AQ15": [{"torque_Nm": 150, "time_s": 25, "kW":0.02,"nom_A":1.3, "start_A":1.6}],
        "AQ25": [{"torque_Nm": 250, "time_s": 25, "kW":0.04,"nom_A":1.5, "start_A":1.7}],
        "AQ30": [{"torque_Nm": 300, "time_s": 30, "kW":0.04,"nom_A":1.5, "start_A":1.7}],
        "AQ50": [{"torque_Nm": 500, "time_s": 30, "kW":0.06,"nom_A":2.5, "start_A":4.2},
                 {"torque_Nm": 500, "time_s": 45, "kW":0.04,"nom_A":1.5, "start_A":1.7}],
        "AQ80": [{"torque_Nm": 800, "time_s": 45, "kW":0.06,"nom_A":2.5, "start_A":4.2}],
        "AQ150":[{"torque_Nm":1140, "time_s": 17, "kW":0.7, "nom_A":8.3, "start_A":33},
                 {"torque_Nm":1500, "time_s": 33, "kW":0.7, "nom_A":8.3, "start_A":33},
                 {"torque_Nm":1500, "time_s": 83, "kW":0.3, "nom_A":4.5, "start_A":11}],
        "AQ280":[{"torque_Nm":2800, "time_s": 58, "kW":0.7, "nom_A":8.3, "start_A":33},
                 {"torque_Nm":2800, "time_s": 83, "kW":0.7, "nom_A":8.3, "start_A":33},
                 {"torque_Nm":2800, "time_s":117, "kW":0.7, "nom_A":8.3, "start_A":33}],
        "AQ430":[{"torque_Nm":4300, "time_s": 33, "kW":1.8, "nom_A":21,  "start_A":77},
                 {"torque_Nm":4300, "time_s": 58, "kW":1.8, "nom_A":21,  "start_A":77},
                 {"torque_Nm":4300, "time_s":100, "kW":0.5, "nom_A":8.5, "start_A":19}],
        "AQ610":[{"torque_Nm":6100, "time_s": 42, "kW":3.0, "nom_A":39,  "start_A":142},
                 {"torque_Nm":6100, "time_s": 83, "kW":1.8, "nom_A":21,  "start_A":77},
                 {"torque_Nm":6100, "time_s":117, "kW":1.8, "nom_A":21,  "start_A":77}],
        "AQ830":[{"torque_Nm":6015, "time_s": 96, "kW":0.7, "nom_A":8.3, "start_A":33},
                 {"torque_Nm":8300, "time_s":192, "kW":0.7, "nom_A":8.3, "start_A":33}],
        "AQ1000":[{"torque_Nm":10400,"time_s":158,"kW":0.7, "nom_A":8.3, "start_A":33},
                  {"torque_Nm":10400,"time_s":112,"kW":1.2, "nom_A":15,  "start_A":61},
                  {"torque_Nm":10400,"time_s": 83,"kW":1.2, "nom_A":15,  "start_A":61},
                  {"torque_Nm":10400,"time_s": 46,"kW":1.8, "nom_A":21,  "start_A":77}],
    },
    # 1x120VAC 60Hz covers AQ5-AQ80 only (larger models not available)
    "1x120VAC_60Hz": {
        "AQ5":  [{"torque_Nm":  50, "time_s": 16, "kW":0.02,"nom_A":1.1, "start_A":1.4}],
        "AQ10": [{"torque_Nm": 100, "time_s": 25, "kW":0.02,"nom_A":1.1, "start_A":1.4}],
        "AQ15": [{"torque_Nm": 150, "time_s": 30, "kW":0.02,"nom_A":1.3, "start_A":1.6}],
        "AQ25": [{"torque_Nm": 250, "time_s": 30, "kW":0.04,"nom_A":1.5, "start_A":1.7}],
        "AQ30": [{"torque_Nm": 300, "time_s": 35, "kW":0.04,"nom_A":1.5, "start_A":1.7}],
        "AQ50": [{"torque_Nm": 500, "time_s": 35, "kW":0.06,"nom_A":2.5, "start_A":4.2},
                 {"torque_Nm": 500, "time_s": 55, "kW":0.04,"nom_A":1.5, "start_A":1.7}],
        "AQ80": [{"torque_Nm": 800, "time_s": 55, "kW":0.06,"nom_A":2.5, "start_A":4.2}],
    },
    "1x120VAC_50Hz": {
        "AQ5":  [{"torque_Nm":  50, "time_s": 13, "kW":0.02,"nom_A":1.1, "start_A":1.4}],
        "AQ10": [{"torque_Nm": 100, "time_s": 21, "kW":0.02,"nom_A":1.1, "start_A":1.4}],
        "AQ15": [{"torque_Nm": 150, "time_s": 25, "kW":0.02,"nom_A":1.3, "start_A":1.6}],
        "AQ25": [{"torque_Nm": 250, "time_s": 25, "kW":0.04,"nom_A":1.5, "start_A":1.7}],
        "AQ30": [{"torque_Nm": 300, "time_s": 30, "kW":0.04,"nom_A":1.5, "start_A":1.7}],
        "AQ50": [{"torque_Nm": 500, "time_s": 30, "kW":0.06,"nom_A":2.5, "start_A":4.2},
                 {"torque_Nm": 500, "time_s": 45, "kW":0.04,"nom_A":1.5, "start_A":1.7}],
        "AQ80": [{"torque_Nm": 800, "time_s": 45, "kW":0.06,"nom_A":2.5, "start_A":4.2}],
        "AQ150":[{"torque_Nm":1140, "time_s": 17, "kW":0.7, "nom_A":8.3, "start_A":33},
                 {"torque_Nm":1500, "time_s": 33, "kW":0.7, "nom_A":8.3, "start_A":33},
                 {"torque_Nm":1500, "time_s": 83, "kW":0.3, "nom_A":4.5, "start_A":11}],
        "AQ280":[{"torque_Nm":2800, "time_s": 58, "kW":0.7, "nom_A":8.3, "start_A":33},
                 {"torque_Nm":2800, "time_s": 83, "kW":0.7, "nom_A":8.3, "start_A":33},
                 {"torque_Nm":2800, "time_s":117, "kW":0.7, "nom_A":8.3, "start_A":33}],
        "AQ430":[{"torque_Nm":4300, "time_s": 33, "kW":1.8, "nom_A":21,  "start_A":77},
                 {"torque_Nm":4300, "time_s": 58, "kW":1.8, "nom_A":21,  "start_A":77},
                 {"torque_Nm":4300, "time_s":100, "kW":0.5, "nom_A":8.5, "start_A":19}],
        "AQ610":[{"torque_Nm":6100, "time_s": 42, "kW":3.0, "nom_A":39,  "start_A":142},
                 {"torque_Nm":6100, "time_s": 83, "kW":1.8, "nom_A":21,  "start_A":77},
                 {"torque_Nm":6100, "time_s":117, "kW":1.8, "nom_A":21,  "start_A":77}],
        "AQ830":[{"torque_Nm":6015, "time_s": 96, "kW":0.7, "nom_A":8.3, "start_A":33},
                 {"torque_Nm":8300, "time_s":192, "kW":0.7, "nom_A":8.3, "start_A":33}],
        "AQ1000":[{"torque_Nm":10400,"time_s":158,"kW":0.7, "nom_A":8.3, "start_A":33},
                  {"torque_Nm":10400,"time_s":112,"kW":1.2, "nom_A":15,  "start_A":61},
                  {"torque_Nm":10400,"time_s": 83,"kW":1.2, "nom_A":15,  "start_A":61},
                  {"torque_Nm":10400,"time_s": 46,"kW":1.8, "nom_A":21,  "start_A":77}],
    },
    "1x220-230VAC_50Hz": {
        "AQ5":  [{"torque_Nm":  50, "time_s": 16, "kW":0.02,"nom_A":0.6, "start_A":0.7}],
        "AQ10": [{"torque_Nm": 100, "time_s": 25, "kW":0.02,"nom_A":0.6, "start_A":0.7}],
        "AQ15": [{"torque_Nm": 150, "time_s": 30, "kW":0.02,"nom_A":0.8, "start_A":1.1}],
        "AQ25": [{"torque_Nm": 250, "time_s": 30, "kW":0.04,"nom_A":1.1, "start_A":1.4}],
        "AQ30": [{"torque_Nm": 300, "time_s": 35, "kW":0.04,"nom_A":1.1, "start_A":1.4}],
        "AQ50": [{"torque_Nm": 500, "time_s": 35, "kW":0.06,"nom_A":1.2, "start_A":1.7},
                 {"torque_Nm": 500, "time_s": 55, "kW":0.04,"nom_A":1.1, "start_A":1.4}],
        "AQ80": [{"torque_Nm": 800, "time_s": 55, "kW":0.06,"nom_A":1.2, "start_A":1.7}],
        "AQ150":[{"torque_Nm":1140, "time_s": 20, "kW":0.7, "nom_A":4.5, "start_A":11},
                 {"torque_Nm":1500, "time_s": 40, "kW":0.7, "nom_A":4.5, "start_A":11},
                 {"torque_Nm":1500, "time_s":100, "kW":0.3, "nom_A":2.1, "start_A":3.9}],
        "AQ280":[{"torque_Nm":2800, "time_s": 70, "kW":0.7, "nom_A":4.5, "start_A":11},
                 {"torque_Nm":2800, "time_s":100, "kW":0.7, "nom_A":4.5, "start_A":11},
                 {"torque_Nm":2800, "time_s":140, "kW":0.7, "nom_A":4.5, "start_A":11}],
        "AQ430":[{"torque_Nm":4300, "time_s": 40, "kW":1.8, "nom_A":11,  "start_A":34},
                 {"torque_Nm":4300, "time_s": 70, "kW":1.8, "nom_A":11,  "start_A":34},
                 {"torque_Nm":4300, "time_s":120, "kW":0.5, "nom_A":3.5, "start_A":7.7}],
        "AQ610":[{"torque_Nm":6100, "time_s": 50, "kW":3.0, "nom_A":19,  "start_A":55},
                 {"torque_Nm":6100, "time_s":100, "kW":1.8, "nom_A":11,  "start_A":34},
                 {"torque_Nm":6100, "time_s":140, "kW":1.8, "nom_A":11,  "start_A":34}],
        "AQ830":[{"torque_Nm":6015, "time_s":115, "kW":0.7, "nom_A":4.5, "start_A":11},
                 {"torque_Nm":8300, "time_s":230, "kW":0.7, "nom_A":4.5, "start_A":11}],
        "AQ1000":[{"torque_Nm":10400,"time_s":190,"kW":0.7, "nom_A":4.5, "start_A":11},
                  {"torque_Nm":10400,"time_s":135,"kW":1.2, "nom_A":8.0, "start_A":21},
                  {"torque_Nm":10400,"time_s":100,"kW":1.2, "nom_A":8.0, "start_A":21},
                  {"torque_Nm":10400,"time_s": 56,"kW":1.8, "nom_A":11,  "start_A":34}],
    },
    "1x220-230VAC_60Hz": {
        "AQ5":  [{"torque_Nm":  50, "time_s": 13, "kW":0.02,"nom_A":0.6, "start_A":0.7}],
        "AQ10": [{"torque_Nm": 100, "time_s": 21, "kW":0.02,"nom_A":0.6, "start_A":0.7}],
        "AQ15": [{"torque_Nm": 150, "time_s": 25, "kW":0.02,"nom_A":0.8, "start_A":1.1}],
        "AQ25": [{"torque_Nm": 250, "time_s": 25, "kW":0.04,"nom_A":1.1, "start_A":1.4}],
        "AQ30": [{"torque_Nm": 300, "time_s": 30, "kW":0.04,"nom_A":1.1, "start_A":1.4}],
        "AQ50": [{"torque_Nm": 500, "time_s": 30, "kW":0.06,"nom_A":1.2, "start_A":1.7},
                 {"torque_Nm": 500, "time_s": 45, "kW":0.04,"nom_A":1.1, "start_A":1.4}],
        "AQ80": [{"torque_Nm": 800, "time_s": 45, "kW":0.06,"nom_A":1.2, "start_A":1.7}],
        "AQ150":[{"torque_Nm":1140, "time_s": 17, "kW":0.7, "nom_A":4.5, "start_A":11},
                 {"torque_Nm":1500, "time_s": 33, "kW":0.7, "nom_A":4.5, "start_A":11},
                 {"torque_Nm":1500, "time_s": 83, "kW":0.3, "nom_A":2.1, "start_A":3.9}],
        "AQ280":[{"torque_Nm":2800, "time_s": 58, "kW":0.7, "nom_A":4.5, "start_A":11},
                 {"torque_Nm":2800, "time_s": 83, "kW":0.7, "nom_A":4.5, "start_A":11},
                 {"torque_Nm":2800, "time_s":117, "kW":0.7, "nom_A":4.5, "start_A":11}],
        "AQ430":[{"torque_Nm":4300, "time_s": 33, "kW":1.8, "nom_A":11,  "start_A":34},
                 {"torque_Nm":4300, "time_s": 58, "kW":1.8, "nom_A":11,  "start_A":34},
                 {"torque_Nm":4300, "time_s":100, "kW":0.5, "nom_A":3.5, "start_A":7.7}],
        "AQ610":[{"torque_Nm":6100, "time_s": 42, "kW":3.0, "nom_A":19,  "start_A":55},
                 {"torque_Nm":6100, "time_s": 83, "kW":1.8, "nom_A":11,  "start_A":34},
                 {"torque_Nm":6100, "time_s":117, "kW":1.8, "nom_A":11,  "start_A":34}],
        "AQ830":[{"torque_Nm":6015, "time_s": 96, "kW":0.7, "nom_A":4.5, "start_A":11},
                 {"torque_Nm":8300, "time_s":192, "kW":0.7, "nom_A":4.5, "start_A":11}],
        "AQ1000":[{"torque_Nm":10400,"time_s":158,"kW":0.7, "nom_A":4.5, "start_A":11},
                  {"torque_Nm":10400,"time_s":112,"kW":1.2, "nom_A":8.0, "start_A":21},
                  {"torque_Nm":10400,"time_s": 83,"kW":1.2, "nom_A":8.0, "start_A":21},
                  {"torque_Nm":10400,"time_s": 46,"kW":1.8, "nom_A":11,  "start_A":34}],
    },
    "1x240VAC_50Hz": {
        "AQ5":  [{"torque_Nm":  50, "time_s": 16, "kW":0.02,"nom_A":0.6, "start_A":0.7}],
        "AQ10": [{"torque_Nm": 100, "time_s": 25, "kW":0.02,"nom_A":0.6, "start_A":0.7}],
        "AQ15": [{"torque_Nm": 150, "time_s": 30, "kW":0.02,"nom_A":0.8, "start_A":1.1}],
        "AQ25": [{"torque_Nm": 250, "time_s": 30, "kW":0.04,"nom_A":1.1, "start_A":1.4}],
        "AQ30": [{"torque_Nm": 300, "time_s": 35, "kW":0.04,"nom_A":1.1, "start_A":1.4}],
        "AQ50": [{"torque_Nm": 500, "time_s": 35, "kW":0.06,"nom_A":1.2, "start_A":1.7},
                 {"torque_Nm": 500, "time_s": 55, "kW":0.04,"nom_A":1.1, "start_A":1.4}],
        "AQ80": [{"torque_Nm": 800, "time_s": 55, "kW":0.06,"nom_A":1.2, "start_A":1.7}],
        "AQ150":[{"torque_Nm":1140, "time_s": 20, "kW":0.7, "nom_A":4.5, "start_A":11},
                 {"torque_Nm":1500, "time_s": 40, "kW":0.7, "nom_A":4.5, "start_A":11},
                 {"torque_Nm":1500, "time_s":100, "kW":0.3, "nom_A":2.1, "start_A":3.9}],
        "AQ280":[{"torque_Nm":2800, "time_s": 70, "kW":0.7, "nom_A":4.5, "start_A":11},
                 {"torque_Nm":2800, "time_s":100, "kW":0.7, "nom_A":4.5, "start_A":11},
                 {"torque_Nm":2800, "time_s":140, "kW":0.7, "nom_A":4.5, "start_A":11}],
        "AQ430":[{"torque_Nm":4300, "time_s": 40, "kW":1.8, "nom_A":11,  "start_A":34},
                 {"torque_Nm":4300, "time_s": 70, "kW":1.8, "nom_A":11,  "start_A":34},
                 {"torque_Nm":4300, "time_s":120, "kW":0.5, "nom_A":3.5, "start_A":7.7}],
        "AQ610":[{"torque_Nm":6100, "time_s": 50, "kW":3.0, "nom_A":19,  "start_A":55},
                 {"torque_Nm":6100, "time_s":100, "kW":1.8, "nom_A":11,  "start_A":34},
                 {"torque_Nm":6100, "time_s":140, "kW":1.8, "nom_A":11,  "start_A":34}],
        "AQ830":[{"torque_Nm":6015, "time_s":115, "kW":0.7, "nom_A":4.5, "start_A":11},
                 {"torque_Nm":8300, "time_s":230, "kW":0.7, "nom_A":4.5, "start_A":11}],
        "AQ1000":[{"torque_Nm":10400,"time_s":190,"kW":0.7, "nom_A":4.5, "start_A":11},
                  {"torque_Nm":10400,"time_s":135,"kW":1.2, "nom_A":8.0, "start_A":21},
                  {"torque_Nm":10400,"time_s":100,"kW":1.2, "nom_A":8.0, "start_A":21},
                  {"torque_Nm":10400,"time_s": 56,"kW":1.8, "nom_A":11,  "start_A":34}],
    },
    "3x380VAC_50Hz": {
        "AQ5":  [{"torque_Nm":  50, "time_s": 16, "kW":0.02,"nom_A":0.15,"start_A":0.4}],
        "AQ10": [{"torque_Nm": 100, "time_s": 25, "kW":0.02,"nom_A":0.15,"start_A":0.4}],
        "AQ15": [{"torque_Nm": 150, "time_s": 30, "kW":0.02,"nom_A":0.15,"start_A":0.4}],
        "AQ25": [{"torque_Nm": 250, "time_s": 30, "kW":0.04,"nom_A":0.2, "start_A":0.4}],
        "AQ30": [{"torque_Nm": 300, "time_s": 35, "kW":0.04,"nom_A":0.2, "start_A":0.4}],
        "AQ50": [{"torque_Nm": 500, "time_s": 35, "kW":0.06,"nom_A":0.4, "start_A":0.7},
                 {"torque_Nm": 500, "time_s": 55, "kW":0.04,"nom_A":0.2, "start_A":0.4}],
        "AQ80": [{"torque_Nm": 800, "time_s": 55, "kW":0.06,"nom_A":0.4, "start_A":0.7}],
        "AQ150":[{"torque_Nm":1500, "time_s": 20, "kW":0.73,"nom_A":1.8, "start_A":6.3},
                 {"torque_Nm":1500, "time_s": 40, "kW":0.37,"nom_A":0.92,"start_A":3.6},
                 {"torque_Nm":1500, "time_s":100, "kW":0.14,"nom_A":0.66,"start_A":2.0}],
        "AQ280":[{"torque_Nm":2800, "time_s": 70, "kW":0.73,"nom_A":1.8, "start_A":6.3},
                 {"torque_Nm":2800, "time_s":100, "kW":0.37,"nom_A":0.92,"start_A":3.6},
                 {"torque_Nm":2800, "time_s":140, "kW":0.37,"nom_A":0.92,"start_A":3.6}],
        "AQ430":[{"torque_Nm":4300, "time_s": 40, "kW":0.82,"nom_A":2.6, "start_A":14},
                 {"torque_Nm":4300, "time_s": 70, "kW":0.82,"nom_A":2.6, "start_A":14},
                 {"torque_Nm":4300, "time_s":120, "kW":0.29,"nom_A":0.98,"start_A":3.0}],
        "AQ610":[{"torque_Nm":6100, "time_s": 50, "kW":1.4, "nom_A":3.3, "start_A":23},
                 {"torque_Nm":6100, "time_s":100, "kW":0.82,"nom_A":2.6, "start_A":14},
                 {"torque_Nm":6100, "time_s":140, "kW":0.73,"nom_A":1.8, "start_A":6.3}],
        "AQ830":[{"torque_Nm":8300, "time_s":115, "kW":0.73,"nom_A":1.8, "start_A":6.3},
                 {"torque_Nm":8300, "time_s":230, "kW":0.37,"nom_A":0.92,"start_A":3.6}],
        "AQ1000":[{"torque_Nm":10400,"time_s":190,"kW":0.37,"nom_A":0.92,"start_A":3.6},
                  {"torque_Nm":10400,"time_s": 56,"kW":1.9, "nom_A":4.7, "start_A":20},
                  {"torque_Nm":10400,"time_s":100,"kW":0.87,"nom_A":1.8, "start_A":10},
                  {"torque_Nm":10400,"time_s":135,"kW":0.46,"nom_A":1.6, "start_A":4.8}],
    },
    "3x380VAC_60Hz": {
        "AQ5":  [{"torque_Nm":  50, "time_s": 13, "kW":0.02,"nom_A":0.13,"start_A":0.33}],
        "AQ10": [{"torque_Nm":  90, "time_s": 21, "kW":0.02,"nom_A":0.13,"start_A":0.33}],
        "AQ15": [{"torque_Nm": 100, "time_s": 25, "kW":0.02,"nom_A":0.13,"start_A":0.33}],
        "AQ25": [{"torque_Nm": 250, "time_s": 25, "kW":0.03,"nom_A":0.17,"start_A":0.33}],
        "AQ30": [{"torque_Nm": 300, "time_s": 30, "kW":0.03,"nom_A":0.17,"start_A":0.33}],
        "AQ50": [{"torque_Nm": 500, "time_s": 30, "kW":0.05,"nom_A":0.33,"start_A":0.58},
                 {"torque_Nm": 500, "time_s": 45, "kW":0.03,"nom_A":0.17,"start_A":0.33}],
        "AQ80": [{"torque_Nm": 800, "time_s": 45, "kW":0.05,"nom_A":0.33,"start_A":0.58}],
        "AQ150":[{"torque_Nm":1190, "time_s": 17, "kW":0.61,"nom_A":1.5, "start_A":5.2},
                 {"torque_Nm":1459, "time_s": 33, "kW":0.31,"nom_A":0.76,"start_A":3.0},
                 {"torque_Nm":1500, "time_s": 83, "kW":0.12,"nom_A":0.55,"start_A":1.7}],
        "AQ280":[{"torque_Nm":2800, "time_s": 58, "kW":0.61,"nom_A":1.5, "start_A":5.2},
                 {"torque_Nm":2800, "time_s": 83, "kW":0.61,"nom_A":1.5, "start_A":5.2},
                 {"torque_Nm":2800, "time_s":117, "kW":0.31,"nom_A":0.76,"start_A":3.0}],
        "AQ430":[{"torque_Nm":4300, "time_s": 33, "kW":1.6, "nom_A":3.9, "start_A":17},
                 {"torque_Nm":4300, "time_s": 58, "kW":0.68,"nom_A":2.1, "start_A":11},
                 {"torque_Nm":4300, "time_s":100, "kW":0.38,"nom_A":1.3, "start_A":3.9}],
        "AQ610":[{"torque_Nm":6100, "time_s": 42, "kW":1.9, "nom_A":3.9, "start_A":26},
                 {"torque_Nm":6100, "time_s": 83, "kW":0.68,"nom_A":2.1, "start_A":11},
                 {"torque_Nm":6100, "time_s":117, "kW":0.68,"nom_A":2.1, "start_A":11}],
        "AQ830":[{"torque_Nm":6283, "time_s": 96, "kW":0.61,"nom_A":1.5, "start_A":5.2},
                 {"torque_Nm":8300, "time_s":192, "kW":0.31,"nom_A":0.76,"start_A":3.0}],
        "AQ1000":[{"torque_Nm":10400,"time_s":158,"kW":0.61,"nom_A":1.5, "start_A":5.2},
                  {"torque_Nm":10400,"time_s": 46,"kW":1.6, "nom_A":3.9, "start_A":17},
                  {"torque_Nm":10400,"time_s": 83,"kW":1.2, "nom_A":2.7, "start_A":19},
                  {"torque_Nm":10400,"time_s":112,"kW":0.38,"nom_A":1.3, "start_A":3.9}],
    },
    # 3x400VAC 50Hz — standard European supply. Values corrected (old dict had 3x415V data).
    "3x400VAC_50Hz": {
        "AQ5":  [{"torque_Nm":  50, "time_s": 16, "kW":0.03,"nom_A":0.16,"start_A":0.43}],
        "AQ10": [{"torque_Nm": 100, "time_s": 25, "kW":0.03,"nom_A":0.16,"start_A":0.43}],
        "AQ15": [{"torque_Nm": 150, "time_s": 30, "kW":0.03,"nom_A":0.16,"start_A":0.43}],
        "AQ25": [{"torque_Nm": 250, "time_s": 30, "kW":0.04,"nom_A":0.22,"start_A":0.43}],
        "AQ30": [{"torque_Nm": 300, "time_s": 35, "kW":0.04,"nom_A":0.22,"start_A":0.43}],
        "AQ50": [{"torque_Nm": 500, "time_s": 35, "kW":0.07,"nom_A":0.43,"start_A":0.75},
                 {"torque_Nm": 500, "time_s": 55, "kW":0.04,"nom_A":0.22,"start_A":0.43}],
        "AQ80": [{"torque_Nm": 800, "time_s": 55, "kW":0.07,"nom_A":0.43,"start_A":0.75}],
        "AQ150":[{"torque_Nm":1500, "time_s": 20, "kW":0.8, "nom_A":1.9, "start_A":6.7},
                 {"torque_Nm":1500, "time_s": 40, "kW":0.4, "nom_A":0.97,"start_A":3.8},
                 {"torque_Nm":1500, "time_s":100, "kW":0.15,"nom_A":0.7, "start_A":2.1}],
        "AQ280":[{"torque_Nm":2800, "time_s": 70, "kW":0.8, "nom_A":1.9, "start_A":6.7},
                 {"torque_Nm":2800, "time_s":100, "kW":0.4, "nom_A":0.97,"start_A":3.8},
                 {"torque_Nm":2800, "time_s":140, "kW":0.4, "nom_A":0.97,"start_A":3.8}],
        "AQ430":[{"torque_Nm":4300, "time_s": 40, "kW":0.9, "nom_A":2.7, "start_A":14},
                 {"torque_Nm":4300, "time_s": 70, "kW":0.9, "nom_A":2.7, "start_A":14},
                 {"torque_Nm":4300, "time_s":120, "kW":0.32,"nom_A":1.1, "start_A":3.2}],
        "AQ610":[{"torque_Nm":6100, "time_s": 50, "kW":1.6, "nom_A":3.1, "start_A":20},
                 {"torque_Nm":6100, "time_s":100, "kW":0.9, "nom_A":2.7, "start_A":14},
                 {"torque_Nm":6100, "time_s":140, "kW":0.8, "nom_A":1.9, "start_A":6.7}],
        "AQ830":[{"torque_Nm":8300, "time_s":115, "kW":0.8, "nom_A":1.9, "start_A":6.7},
                 {"torque_Nm":8300, "time_s":230, "kW":0.4, "nom_A":0.97,"start_A":3.8}],
        "AQ1000":[{"torque_Nm":10400,"time_s":190,"kW":0.4, "nom_A":0.97,"start_A":3.8},
                  {"torque_Nm":10400,"time_s": 56,"kW":1.3, "nom_A":3.2, "start_A":14},
                  {"torque_Nm":10400,"time_s":100,"kW":0.96,"nom_A":1.9, "start_A":11},
                  {"torque_Nm":10400,"time_s":135,"kW":0.5, "nom_A":1.6, "start_A":5.0}],
    },
    "3x400VAC_60Hz": {
        "AQ5":  [{"torque_Nm":  50, "time_s": 13, "kW":0.02,"nom_A":0.13,"start_A":0.35}],
        "AQ10": [{"torque_Nm":  90, "time_s": 21, "kW":0.02,"nom_A":0.13,"start_A":0.35}],
        "AQ15": [{"torque_Nm": 110, "time_s": 25, "kW":0.02,"nom_A":0.13,"start_A":0.35}],
        "AQ25": [{"torque_Nm": 250, "time_s": 25, "kW":0.04,"nom_A":0.18,"start_A":0.35}],
        "AQ30": [{"torque_Nm": 300, "time_s": 30, "kW":0.04,"nom_A":0.18,"start_A":0.35}],
        "AQ50": [{"torque_Nm": 500, "time_s": 30, "kW":0.06,"nom_A":0.35,"start_A":0.61},
                 {"torque_Nm": 500, "time_s": 45, "kW":0.04,"nom_A":0.18,"start_A":0.35}],
        "AQ80": [{"torque_Nm": 800, "time_s": 45, "kW":0.06,"nom_A":0.35,"start_A":0.61}],
        "AQ150":[{"torque_Nm":1319, "time_s": 17, "kW":0.67,"nom_A":1.6, "start_A":5.5},
                 {"torque_Nm":1500, "time_s": 33, "kW":0.34,"nom_A":0.8, "start_A":3.1},
                 {"torque_Nm":1500, "time_s": 83, "kW":0.13,"nom_A":0.58,"start_A":1.8}],
        "AQ280":[{"torque_Nm":2800, "time_s": 58, "kW":0.67,"nom_A":1.6, "start_A":5.5},
                 {"torque_Nm":2800, "time_s": 83, "kW":0.67,"nom_A":1.6, "start_A":5.5},
                 {"torque_Nm":2800, "time_s":117, "kW":0.34,"nom_A":0.8, "start_A":3.1}],
        "AQ430":[{"torque_Nm":4300, "time_s": 33, "kW":1.7, "nom_A":4.1, "start_A":18},
                 {"torque_Nm":4300, "time_s": 58, "kW":0.75,"nom_A":2.3, "start_A":12},
                 {"torque_Nm":4300, "time_s":100, "kW":0.42,"nom_A":1.4, "start_A":4.1}],
        "AQ610":[{"torque_Nm":6100, "time_s": 42, "kW":1.3, "nom_A":2.8, "start_A":20},
                 {"torque_Nm":6100, "time_s": 83, "kW":0.75,"nom_A":2.3, "start_A":12},
                 {"torque_Nm":6100, "time_s":117, "kW":0.75,"nom_A":2.3, "start_A":12}],
        "AQ830":[{"torque_Nm":6961, "time_s": 96, "kW":0.67,"nom_A":1.6, "start_A":5.5},
                 {"torque_Nm":8300, "time_s":192, "kW":0.34,"nom_A":0.8, "start_A":3.1}],
        "AQ1000":[{"torque_Nm":10400,"time_s":158,"kW":0.67,"nom_A":1.6, "start_A":5.5},
                  {"torque_Nm":10400,"time_s": 46,"kW":1.7, "nom_A":4.1, "start_A":18},
                  {"torque_Nm":10400,"time_s": 83,"kW":1.3, "nom_A":2.8, "start_A":20},
                  {"torque_Nm":10400,"time_s":112,"kW":0.42,"nom_A":1.4, "start_A":4.1}],
    },
    "3x415VAC_50Hz": {
        "AQ5":  [{"torque_Nm":  50, "time_s": 16, "kW":0.03,"nom_A":0.17,"start_A":0.45}],
        "AQ10": [{"torque_Nm": 100, "time_s": 25, "kW":0.03,"nom_A":0.17,"start_A":0.45}],
        "AQ15": [{"torque_Nm": 150, "time_s": 30, "kW":0.03,"nom_A":0.17,"start_A":0.45}],
        "AQ25": [{"torque_Nm": 250, "time_s": 30, "kW":0.05,"nom_A":0.23,"start_A":0.45}],
        "AQ30": [{"torque_Nm": 300, "time_s": 35, "kW":0.05,"nom_A":0.23,"start_A":0.45}],
        "AQ50": [{"torque_Nm": 500, "time_s": 35, "kW":0.07,"nom_A":0.45,"start_A":0.79},
                 {"torque_Nm": 500, "time_s": 55, "kW":0.05,"nom_A":0.23,"start_A":0.45}],
        "AQ80": [{"torque_Nm": 800, "time_s": 55, "kW":0.07,"nom_A":0.45,"start_A":0.79}],
        "AQ150":[{"torque_Nm":1500, "time_s": 20, "kW":0.87,"nom_A":2.0, "start_A":7.0},
                 {"torque_Nm":1500, "time_s": 40, "kW":0.28,"nom_A":0.65,"start_A":2.6},
                 {"torque_Nm":1500, "time_s":100, "kW":0.11,"nom_A":0.49,"start_A":1.4}],
        "AQ280":[{"torque_Nm":2800, "time_s": 70, "kW":0.87,"nom_A":2.0, "start_A":7.0},
                 {"torque_Nm":2800, "time_s":100, "kW":0.44,"nom_A":1.1, "start_A":4.0},
                 {"torque_Nm":2800, "time_s":140, "kW":0.28,"nom_A":0.65,"start_A":2.6}],
        "AQ430":[{"torque_Nm":4300, "time_s": 40, "kW":0.97,"nom_A":2.9, "start_A":15},
                 {"torque_Nm":4300, "time_s": 70, "kW":0.97,"nom_A":2.9, "start_A":15},
                 {"torque_Nm":4300, "time_s":120, "kW":0.35,"nom_A":1.1, "start_A":3.3}],
        "AQ610":[{"torque_Nm":6100, "time_s": 50, "kW":1.8, "nom_A":3.2, "start_A":21},
                 {"torque_Nm":6100, "time_s":100, "kW":0.97,"nom_A":2.9, "start_A":15},
                 {"torque_Nm":6100, "time_s":140, "kW":0.87,"nom_A":2.0, "start_A":7.0}],
        "AQ830":[{"torque_Nm":8300, "time_s":115, "kW":0.87,"nom_A":2.0, "start_A":7.0},
                 {"torque_Nm":8300, "time_s":230, "kW":0.28,"nom_A":0.65,"start_A":2.6}],
        "AQ1000":[{"torque_Nm":10400,"time_s":190,"kW":0.44,"nom_A":1.1, "start_A":4.0},
                  {"torque_Nm":10400,"time_s": 56,"kW":1.4, "nom_A":3.3, "start_A":14},
                  {"torque_Nm":10400,"time_s":100,"kW":1.1, "nom_A":1.9, "start_A":11},
                  {"torque_Nm":10400,"time_s":135,"kW":0.54,"nom_A":1.7, "start_A":5.3}],
    },
    "3x440VAC_60Hz": {
        "AQ5":  [{"torque_Nm":  50, "time_s": 13, "kW":0.03,"nom_A":0.15,"start_A":0.39}],
        "AQ10": [{"torque_Nm": 100, "time_s": 21, "kW":0.03,"nom_A":0.15,"start_A":0.39}],
        "AQ15": [{"torque_Nm": 145, "time_s": 25, "kW":0.03,"nom_A":0.15,"start_A":0.39}],
        "AQ25": [{"torque_Nm": 250, "time_s": 25, "kW":0.04,"nom_A":0.2, "start_A":0.39}],
        "AQ30": [{"torque_Nm": 300, "time_s": 30, "kW":0.04,"nom_A":0.2, "start_A":0.39}],
        "AQ50": [{"torque_Nm": 500, "time_s": 30, "kW":0.07,"nom_A":0.39,"start_A":0.68},
                 {"torque_Nm": 500, "time_s": 45, "kW":0.04,"nom_A":0.2, "start_A":0.39}],
        "AQ80": [{"torque_Nm": 800, "time_s": 45, "kW":0.07,"nom_A":0.39,"start_A":0.68}],
        "AQ150":[{"torque_Nm":1500, "time_s": 17, "kW":0.81,"nom_A":1.7, "start_A":6.1},
                 {"torque_Nm":1500, "time_s": 33, "kW":0.41,"nom_A":0.88,"start_A":3.5},
                 {"torque_Nm":1500, "time_s": 83, "kW":0.16,"nom_A":0.64,"start_A":1.9}],
        "AQ280":[{"torque_Nm":2800, "time_s": 58, "kW":0.81,"nom_A":1.7, "start_A":6.1},
                 {"torque_Nm":2800, "time_s": 83, "kW":0.41,"nom_A":0.88,"start_A":3.5},
                 {"torque_Nm":2800, "time_s":117, "kW":0.41,"nom_A":0.88,"start_A":3.5}],
        "AQ430":[{"torque_Nm":4300, "time_s": 33, "kW":1.3, "nom_A":2.9, "start_A":13},
                 {"torque_Nm":4300, "time_s": 58, "kW":0.91,"nom_A":2.5, "start_A":13},
                 {"torque_Nm":4300, "time_s":100, "kW":0.33,"nom_A":0.94,"start_A":2.9}],
        "AQ610":[{"torque_Nm":6100, "time_s": 42, "kW":1.6, "nom_A":3.1, "start_A":22},
                 {"torque_Nm":6100, "time_s": 83, "kW":0.91,"nom_A":2.5, "start_A":13},
                 {"torque_Nm":6100, "time_s":117, "kW":0.91,"nom_A":2.5, "start_A":13}],
        "AQ830":[{"torque_Nm":7980, "time_s": 96, "kW":0.81,"nom_A":1.7, "start_A":6.1},
                 {"torque_Nm":8300, "time_s":192, "kW":0.41,"nom_A":0.88,"start_A":3.5}],
        "AQ1000":[{"torque_Nm":10400,"time_s":158,"kW":0.41,"nom_A":0.88,"start_A":3.5},
                  {"torque_Nm":10400,"time_s": 46,"kW":2.1, "nom_A":4.5, "start_A":20},
                  {"torque_Nm":10400,"time_s": 83,"kW":0.97,"nom_A":1.7, "start_A":9.7},
                  {"torque_Nm":10400,"time_s":112,"kW":0.51,"nom_A":1.5, "start_A":4.6}],
    },
    "3x460VAC_60Hz": {
        "AQ5":  [{"torque_Nm":  50, "time_s": 13, "kW":0.03,"nom_A":0.16,"start_A":0.41}],
        "AQ10": [{"torque_Nm": 100, "time_s": 21, "kW":0.03,"nom_A":0.16,"start_A":0.41}],
        "AQ15": [{"torque_Nm": 150, "time_s": 25, "kW":0.03,"nom_A":0.16,"start_A":0.41}],
        "AQ25": [{"torque_Nm": 250, "time_s": 25, "kW":0.05,"nom_A":0.21,"start_A":0.41}],
        "AQ30": [{"torque_Nm": 300, "time_s": 30, "kW":0.05,"nom_A":0.21,"start_A":0.41}],
        "AQ50": [{"torque_Nm": 500, "time_s": 30, "kW":0.07,"nom_A":0.41,"start_A":0.71},
                 {"torque_Nm": 500, "time_s": 45, "kW":0.05,"nom_A":0.21,"start_A":0.41}],
        "AQ80": [{"torque_Nm": 800, "time_s": 45, "kW":0.07,"nom_A":0.41,"start_A":0.71}],
        "AQ150":[{"torque_Nm":1500, "time_s": 17, "kW":0.89,"nom_A":1.8, "start_A":6.4},
                 {"torque_Nm":1500, "time_s": 33, "kW":0.45,"nom_A":0.93,"start_A":3.6},
                 {"torque_Nm":1500, "time_s": 83, "kW":0.17,"nom_A":0.67,"start_A":2.0}],
        "AQ280":[{"torque_Nm":2800, "time_s": 58, "kW":0.89,"nom_A":1.8, "start_A":6.4},
                 {"torque_Nm":2800, "time_s": 83, "kW":0.45,"nom_A":0.93,"start_A":3.6},
                 {"torque_Nm":2800, "time_s":117, "kW":0.45,"nom_A":0.93,"start_A":3.6}],
        "AQ430":[{"torque_Nm":4300, "time_s": 33, "kW":1.0, "nom_A":2.6, "start_A":14},
                 {"torque_Nm":4300, "time_s": 58, "kW":1.0, "nom_A":2.6, "start_A":14},
                 {"torque_Nm":4300, "time_s":100, "kW":0.36,"nom_A":0.98,"start_A":3.1}],
        "AQ610":[{"torque_Nm":6100, "time_s": 42, "kW":1.8, "nom_A":3.0, "start_A":19},
                 {"torque_Nm":6100, "time_s": 83, "kW":1.0, "nom_A":2.6, "start_A":14},
                 {"torque_Nm":6100, "time_s":117, "kW":0.89,"nom_A":1.8, "start_A":6.4}],
        "AQ830":[{"torque_Nm":8300, "time_s": 96, "kW":0.89,"nom_A":1.8, "start_A":6.4},
                 {"torque_Nm":8300, "time_s":192, "kW":0.45,"nom_A":0.93,"start_A":3.6}],
        "AQ1000":[{"torque_Nm":10400,"time_s":158,"kW":0.45,"nom_A":0.93,"start_A":3.6},
                  {"torque_Nm":10400,"time_s": 46,"kW":1.0, "nom_A":2.6, "start_A":14},
                  {"torque_Nm":10400,"time_s": 83,"kW":1.1, "nom_A":1.8, "start_A":11},
                  {"torque_Nm":10400,"time_s":112,"kW":0.56,"nom_A":1.6, "start_A":4.8}],
    },
    "3x480VAC_60Hz": {
        "AQ5":  [{"torque_Nm":  50, "time_s": 13, "kW":0.03,"nom_A":0.16,"start_A":0.43}],
        "AQ10": [{"torque_Nm": 100, "time_s": 21, "kW":0.03,"nom_A":0.16,"start_A":0.43}],
        "AQ15": [{"torque_Nm": 150, "time_s": 25, "kW":0.03,"nom_A":0.16,"start_A":0.43}],
        "AQ25": [{"torque_Nm": 250, "time_s": 25, "kW":0.05,"nom_A":0.22,"start_A":0.43}],
        "AQ30": [{"torque_Nm": 300, "time_s": 30, "kW":0.05,"nom_A":0.22,"start_A":0.43}],
        "AQ50": [{"torque_Nm": 500, "time_s": 30, "kW":0.08,"nom_A":0.43,"start_A":0.75},
                 {"torque_Nm": 500, "time_s": 45, "kW":0.05,"nom_A":0.22,"start_A":0.43}],
        "AQ80": [{"torque_Nm": 800, "time_s": 45, "kW":0.08,"nom_A":0.43,"start_A":0.75}],
        "AQ150":[{"torque_Nm":1500, "time_s": 17, "kW":0.96,"nom_A":1.9, "start_A":6.7},
                 {"torque_Nm":1490, "time_s": 33, "kW":0.31,"nom_A":0.62,"start_A":2.5},
                 {"torque_Nm":1500, "time_s": 83, "kW":0.12,"nom_A":0.48,"start_A":1.4}],
        "AQ280":[{"torque_Nm":2800, "time_s": 58, "kW":0.96,"nom_A":1.9, "start_A":6.7},
                 {"torque_Nm":2800, "time_s": 83, "kW":0.48,"nom_A":0.97,"start_A":3.8},
                 {"torque_Nm":2800, "time_s":117, "kW":0.31,"nom_A":0.62,"start_A":2.5}],
        "AQ430":[{"torque_Nm":4300, "time_s": 33, "kW":1.1, "nom_A":2.7, "start_A":14},
                 {"torque_Nm":4300, "time_s": 58, "kW":1.1, "nom_A":2.7, "start_A":14},
                 {"torque_Nm":4300, "time_s":100, "kW":0.39,"nom_A":1.1, "start_A":3.2}],
        "AQ610":[{"torque_Nm":6100, "time_s": 42, "kW":2.0, "nom_A":3.1, "start_A":20},
                 {"torque_Nm":6100, "time_s": 83, "kW":1.1, "nom_A":2.7, "start_A":14},
                 {"torque_Nm":6100, "time_s":117, "kW":0.96,"nom_A":1.9, "start_A":6.7}],
        "AQ830":[{"torque_Nm":8300, "time_s": 96, "kW":0.96,"nom_A":1.9, "start_A":6.7},
                 {"torque_Nm":8300, "time_s":192, "kW":0.31,"nom_A":0.62,"start_A":2.5}],
        "AQ1000":[{"torque_Nm":10400,"time_s":158,"kW":0.48,"nom_A":0.97,"start_A":3.8},
                  {"torque_Nm":10400,"time_s": 46,"kW":1.1, "nom_A":2.7, "start_A":14},
                  {"torque_Nm":10400,"time_s": 83,"kW":1.2, "nom_A":1.9, "start_A":11},
                  {"torque_Nm":10400,"time_s":112,"kW":0.6, "nom_A":1.6, "start_A":5.0}],
    },
    # 24VDC available for AQ5-AQ80 only
    "24VDC": {
        "AQ5":  [{"torque_Nm":  50, "time_s": 13, "kW":0.03,"nom_A":0.97,"start_A":8}],
        "AQ10": [{"torque_Nm": 100, "time_s": 21, "kW":0.03,"nom_A":1.18,"start_A":8}],
        "AQ15": [{"torque_Nm": 150, "time_s": 25, "kW":0.03,"nom_A":1.46,"start_A":8}],
        "AQ25": [{"torque_Nm": 250, "time_s": 25, "kW":0.05,"nom_A":1.74,"start_A":10}],
        "AQ30": [{"torque_Nm": 300, "time_s": 30, "kW":0.05,"nom_A":1.98,"start_A":10}],
        "AQ50": [{"torque_Nm": 500, "time_s": 35, "kW":0.05,"nom_A":1.8, "start_A":10}],
        "AQ80": [{"torque_Nm": 800, "time_s": 40, "kW":0.05,"nom_A":2.88,"start_A":10}],
    },
}
# Backwards-compatible alias (old name had 3x415V data — now points to correct 3x400V table)
AQ_PERFORMANCE_3x400V_50HZ = {m: v[0] for m, v in AQ_PERFORMANCE["3x400VAC_50Hz"].items()}

# ── LOGIC variant terminal map (sheets 3.7 / 3.8) ──────────────────────────
# Source: TEC01-03_E+F_GRP_rev06B sheet 3.7 (single-phase overview) and
# sheet 3.8 (3-phase overview). The positioner board (P2) and RS4 board (P1)
# are option cards; their terminals only exist when fitted.
AQ_LOGIC_TERMINALS = {
    "source": "Bernard TEC01-03 rev06B sheets 3.7-3.8 (LOGIC overview)",
    "motor_supply": {
        "1ph": {"L": "L2/L", "N": "L1/N", "PE": "PE"},
        "3ph": {"L1": "L1/N", "L2": "L2/L", "L3": "L3", "PE": "PE",
                "note": "Phase sequence not important; missing phase triggers fault"},
    },
    "remote_commands": {
        3:  "OPEN command (digital input; connect to 10/11 for dry-contact control)",
        4:  "CLOSE command",
        5:  "STOP command",
        6:  "AUX1 configurable command",
        7:  "AUX2 / cancel self-hold (do not connect if self-holding not required)",
        8:  "24 VDC output supply (for powering digital inputs from terminals 3-7)",
        9:  "Emergency 24 VDC input (external 24V keeps actuator alive on power loss)",
        10: "0 V reference (common for digital inputs)",
        11: "0 V reference (common)",
    },
    "signalling_relays": {
        "R1": {"terminals": [20, 21, 22], "function": "Valve OPEN (SPDT)"},
        "R2": {"terminals": [23, 24, 25], "function": "Valve CLOSED (SPDT)"},
        "R3": {"terminals": "configurable", "function": "Configurable relay (SPDT)"},
        "RD": {"terminals": [26, 27, 28],
               "function": "Fault relay — 26-28 CLOSED = actuator available (NC healthy)"},
    },
    "positioner_board_P2": {
        "note": "Optional board; terminals 12-19 only present when fitted",
        12: "Position xmitter supply − (or signal − for 2-wire 4-20mA)",
        13: "Position xmitter signal + (0/4-20mA output)",
        14: "Position xmitter supply + (12-32 VDC external)",
        15: "Position xmitter supply − (3-wire connection only)",
        16: "Torque xmitter output +",
        17: "Torque xmitter output −",
        18: "Analog setpoint input −",
        19: "Analog setpoint input +",
    },
    "rs4_board_P1": {
        "note": "Optional RS4 configurable relays board; terminals 29-36 only when fitted",
        "R4": [29, 30],
        "R5": [31, 32],
        "R6": [33, 34],
        "R7": [35, 36],
    },
}

# ── Wiring sheet index (all 40 sheets) ─────────────────────────────────────
AQ_WIRING_SHEETS = {
    "3.1":  "AQ SWITCH: Single-phase (AQ5-80)",
    "3.2":  "AQ SWITCH: 3-phases (AQ5-80; layout reused on sheet 3.21 for AQ150-1000)",
    "3.3":  "AQ5-15 SWITCH: Single-phase + Positioner option",
    "3.4":  "AQ5-15 SWITCH: 3-phases + Positioner option",
    "3.5":  "AQ25-80 SWITCH: Single-phase + Positioner option",
    "3.6":  "AQ25-80 SWITCH: 3-phases + Positioner option",
    "3.7":  "AQ LOGIC: Single-phase overview (customer terminals + signalling, 2 pages)",
    "3.8":  "AQ LOGIC: 3-phases overview (customer terminals + signalling, 2 pages)",
    "3.9":  "AQ5-15 LOGIC: On-Off",
    "3.10": "AQ5-15 LOGIC: Positioner (2 pages)",
    "3.11": "AQ5-15 LOGIC: Positioner + RS4 (2 pages)",
    "3.12": "AQ5-15 LOGIC: RS4 only (2 pages)",
    "3.13": "AQ5-15 LOGIC: Transmitter (2 pages)",
    "3.14": "AQ5-15 LOGIC: Transmitter + RS4 (2 pages)",
    "3.15": "AQ25-80 LOGIC: On-Off",
    "3.16": "AQ25-80 LOGIC: Positioner (2 pages)",
    "3.17": "AQ25-80 LOGIC: Positioner + RS4 (2 pages)",
    "3.18": "AQ25-80 LOGIC: RS4 only (2 pages)",
    "3.19": "AQ25-80 LOGIC: Transmitter (2 pages)",
    "3.20": "AQ25-80 LOGIC: Transmitter + RS4 (2 pages)",
    "3.21": "AQ150-1000 SWITCH: 3-phases standard",
    "3.22": "AQ150-1000 LOGIC: Single-phase On-Off",
    "3.23": "AQ150-1000 LOGIC: Single-phase (emergency supply)",
    "3.24": "AQ150-1000 LOGIC: 3-phases On-Off",
    "3.25": "AQ150-1000 LOGIC: Single-phase RS4 / 3-phases RS4",
    "3.26": "AQ150-1000 LOGIC: 3-phases RS4",
    "3.27": "AQ150-1000 LOGIC: Single-phase Positioner",
    "3.28": "AQ150-1000 LOGIC: 3-phases Positioner",
    "3.29": "AQ150-1000 LOGIC: Single-phase Transmitter",
    "3.30": "AQ150-1000 LOGIC: 3-phases Transmitter",
    "3.31": "AQ150-1000 LOGIC: Single-phase Positioner + RS4",
    "3.32": "AQ150-1000 LOGIC: 3-phases Positioner + RS4",
    "3.33": "AQ150-1000 LOGIC: Single-phase Transmitter + RS4",
    "3.34": "AQ150-1000 LOGIC: 3-phases Transmitter + RS4",
    "3.35": "AQ5-15 Direct Current SWITCH",
    "3.36": "AQ5-15 Direct Current LOGIC: Positioner + RS4",
    "3.37": "AQ5-15 Direct Current LOGIC: Transmitter + RS4",
    "3.38": "AQ25-80 Direct Current SWITCH",
    "3.39": "AQ25-80 Direct Current LOGIC: Positioner + RS4",
    "3.40": "AQ25-80 Direct Current LOGIC: Transmitter + RS4",
}


@mcp.tool()
def aq_terminals(phase: str = "3ph", variant: str = "SWITCH") -> str:
    """Return Bernard AQ terminal map. variant='SWITCH' returns mechanical switch
    terminal block (sheet 3.1/3.2). variant='LOGIC' returns the LOGIC controller
    customer terminal map (sheets 3.7/3.8). phase='3ph' or '1ph'."""
    import json
    if variant.upper() == "LOGIC":
        return json.dumps(AQ_LOGIC_TERMINALS, indent=2, ensure_ascii=False)
    data = AQ_SWITCH_TERMINALS if phase.lower().startswith("3") else AQ_SWITCH_TERMINALS_1PH
    return json.dumps(data, indent=2, ensure_ascii=False)


@mcp.tool()
def aq_model_data(size: str = "AQ25", voltage: str = "3x400VAC_50Hz") -> str:
    """Return mechanical dimensions + electrical performance for an AQ model.
    size: e.g. 'AQ150'. voltage: key from AQ_PERFORMANCE (default 3x400VAC_50Hz).
    Use aq_performance() to list all available voltages."""
    import json
    size = size.upper().replace(" ", "")
    perf_table = AQ_PERFORMANCE.get(voltage, AQ_PERFORMANCE.get(voltage.replace(" ","_"), {}))
    out = {
        "model": size,
        "mechanical": AQ_DIMENSIONS.get(size, "unknown model"),
        "performance": {
            "voltage": voltage,
            "speeds": perf_table.get(size, "not available at this voltage"),
        },
        "wiring_sheets": AQ_WIRING_SHEETS,
    }
    return json.dumps(out, indent=2, ensure_ascii=False)


@mcp.tool()
def aq_performance(model: str = "AQ150", voltage: str = "") -> str:
    """Return performance data for an AQ model across all voltages, or for a
    specific voltage. model: e.g. 'AQ150'. voltage: optional key like
    '3x400VAC_50Hz' — omit to get all voltages."""
    import json
    model = model.upper().replace(" ", "")
    if voltage:
        table = AQ_PERFORMANCE.get(voltage, {})
        return json.dumps({
            "model": model, "voltage": voltage,
            "speeds": table.get(model, "not available"),
        }, indent=2)
    result = {"model": model, "available_voltages": list(AQ_PERFORMANCE.keys()), "by_voltage": {}}
    for v, table in AQ_PERFORMANCE.items():
        if model in table:
            result["by_voltage"][v] = table[model]
    return json.dumps(result, indent=2)


@mcp.tool()
def draw_aq_valve_control(origin_x: float = 0.0, origin_y: float = 0.0,
                          size: str = "AQ25", tag_prefix: str = "") -> str:
    """Draw the Bernard AQ 3x400VAC motor-operated valve control schematic
    (power + PLC/HMI control sheet), wired to the AQ SWITCH datasheet interface
    (SUG_17003 Ch.11.1, without positioner): 3-phase motor, S1/S2 travel-limit
    switches, open/close torque switches, NC motor thermostat, 230V heater and
    position-signalling switches. The travel-limit, torque and thermostat
    contacts are wired into the contactor coil rungs (hardware stop) AND the
    signalling/status contacts go to the PLC. Reversing pair -KM1(OPEN)/-KM2
    (CLOSE) swaps two phases (SWITCH type, no auto phase correction); -K0 3UG
    phase monitor permissive; -Q2 3RV backup breaker; S7-1200 (1214C) 8DI/2DO;
    KTP700 HMI over PROFINET. `origin_x/_y` shift the sheet, `size` sets the
    actuator label, `tag_prefix` prepends device tags. EFF_* library auto-loads.
    TERMINAL NUMBERS are baked in from Bernard datasheet TEC01-03 sheet 3.2
    (AQ SWITCH, 3-phase): motor 1/2/3 +PE, thermostat 40/41 (NC), torque OPEN
    4-5-6 / CLOSE 7-8-9, travel OPEN 10-11-12 / CLOSE 13-14-15, aux travel
    20-25, heater 26/27, potentiometer 16-18, 4-20mA 80(+)/81(-). The only
    open item is the NO/NC polarity WITHIN each 3-wire switch group, labelled
    'TBC' -- confirm against the order-specific sheet inside the unit cover."""
    import math
    acad = _get_acad()
    doc = acad.ActiveDocument

    # --- layer + symbol library ------------------------------------------------
    try:
        try:
            lay = doc.Layers.Item("EFF-VALVE")
        except Exception:
            lay = doc.Layers.Add("EFF-VALVE")
            lay.color = 7
        doc.ActiveLayer = lay
    except Exception:
        pass
    for nm, builder in _SYMBOLS.items():
        if not _block_exists(doc, nm):
            try:
                builder(doc.Blocks.Add(_point(0, 0), nm))
            except Exception:
                pass

    ms = doc.ModelSpace
    ox, oy = float(origin_x), float(origin_y)
    pre = str(tag_prefix)

    # --- local drawing helpers (apply the origin offset) -----------------------
    def block(name, x, y, s=1.0, rot=0.0):
        ref = ms.InsertBlock(_point(x + ox, y + oy), name,
                             float(s), float(s), float(s), math.radians(rot))
        try:
            ref.Layer = "EFF-VALVE"
        except Exception:
            pass
        return ref

    def line(x1, y1, x2, y2):
        ms.AddLine(_point(x1 + ox, y1 + oy), _point(x2 + ox, y2 + oy))

    def poly(pts):
        ms.AddLightWeightPolyline(_coords([(p[0] + ox, p[1] + oy) for p in pts]))

    def rect(x1, y1, x2, y2):
        p = ms.AddLightWeightPolyline(_coords([(x1 + ox, y1 + oy), (x2 + ox, y1 + oy),
                                               (x2 + ox, y2 + oy), (x1 + ox, y2 + oy)]))
        p.Closed = True

    def text(t, x, y, h, a="left"):
        _add_text(ms, t, x + ox, y + oy, h, a)

    def mtext(t, x, y, w, h):
        try:
            m = ms.AddMText(_point(x + ox, y + oy), float(w),
                            str(t).replace("\\n", "\n").replace("\n", "\\P"))
            try:
                m.Height = float(h)
            except Exception:
                pass
        except Exception:
            pass

    def dot(x, y, r=1.5):
        _add_dot(ms, x + ox, y + oy, r)

    def contact(x, y, kind="nc", w=8.0):
        """Small inline control contact centred at (x, y). kind 'nc' adds the NC bar."""
        line(x - w / 2, y, x - 2, y)            # left lead
        line(x + 2, y, x + w / 2, y)            # right lead
        line(x - 2, y, x + 2, y + 4)            # moving blade
        if kind == "nc":
            line(x + 2, y + 5, x + 2, y + 1)    # NC bar tick

    L1, L2, L3 = 100.0, 108.6, 117.2          # 3-phase bus columns
    K1 = (160.0, 168.6, 177.2)                # KM2 pole columns
    N = 126.0                                 # neutral / PE reference column

    # =====================================================================
    # POWER SHEET (left)
    # =====================================================================
    text(f"VALVE ACTUATOR  -  BERNARD {size} (SWITCH)  3x400VAC  -  POWER CIRCUIT", 40, 600, 7)
    for lbl, x in (("L1", L1), ("L2", L2), ("L3", L3), ("N", N), ("PE", N + 8)):
        text(lbl, x, 586, 4, "center")
    text("3 x 400 VAC + N + PE   50 Hz", 70, 594, 4)

    block("EFF_MCB_3P", L1, 555)              # -Q1 supply MCB
    block("EFF_PHASE_MONITOR", 210, 500)      # -K0 phase monitor
    block("EFF_CONTACTOR_3P", L1, 470)        # -KM1 OPEN  (CCW / S1)
    block("EFF_CONTACTOR_3P", K1[0], 470)     # -KM2 CLOSE (CW / S2)
    block("EFF_MOTORPROT_3P", L1, 400)        # -Q2 motor protection
    block("EFF_MOTOR_3PH", L1, 335)           # actuator motor (E)

    text(f"{pre}-Q1  5SL6316-7  (3P)", 66, 557, 4.5)
    text(f"{pre}-K0  3UG4512  PHASE OK", 232, 507, 3.6)
    text(f"{pre}-KM1  OPEN (CCW/S1)  3TG1010", 60, 485, 4)
    text(f"{pre}-KM2  CLOSE (CW/S2)  3TG1010", K1[0], 485, 4)
    text(f"{pre}-Q2  3RV2011 (+3RV2901 aux)", 36, 384, 4)
    text("U", L1, 320, 3.5, "center")
    text("V", L2, 320, 3.5, "center")
    text("W", L3, 320, 3.5, "center")
    text("ACTUATOR 3~ MOTOR (E)", 60, 312, 3.2)

    for L in (L1, L2, L3):
        line(L, 580, L, 555)
        line(L, 525, L, 470)
    line(N, 580, N, 250)
    line(N + 8, 580, N + 8, 250)
    text("PE", N + 11, 250, 3)

    mon = (210.0, 218.6, 227.2)
    for (L, mx, ty) in ((L1, mon[0], 520), (L2, mon[1], 516), (L3, mon[2], 512)):
        poly([(L, ty), (mx, ty), (mx, 507)])
        dot(L, ty)

    for (L, kx, ty) in ((L1, K1[0], 495), (L2, K1[1], 491), (L3, K1[2], 487)):
        poly([(L, ty), (kx, ty), (kx, 470)])
        dot(L, ty)

    for L in (L1, L2, L3):
        line(L, 440, L, 400)
    for (kx, dest, ty) in ((K1[0], L3, 432), (K1[1], L2, 428), (K1[2], L1, 424)):
        poly([(kx, 440), (kx, ty), (dest, ty)])
        dot(dest, ty)
    for L in (L1, L2, L3):
        line(L, 365, L, 335)

    for (cols, dy, nums) in (((L1, L2, L3), 472, ("1", "3", "5")),
                             ((L1, L2, L3), 433, ("2", "4", "6")),
                             (K1, 472, ("1", "3", "5")),
                             (K1, 433, ("2", "4", "6"))):
        for c, n in zip(cols, nums):
            text(n, c, dy, 3, "center")

    # heater circuit: 230 VAC, permanently energised
    text(f"{pre}-Q3  5SL6216-7 (2P)", 250, 300, 3.6)
    block("EFF_MCB_2P", 250, 290)
    line(L1, 560, 250, 560)
    line(250, 560, 250, 320)
    dot(L1, 560)
    line(N, 360, 258.6, 360)
    line(258.6, 360, 258.6, 320)
    dot(N, 360)
    line(250, 260, 250, 240)
    line(258.6, 260, 258.6, 240)
    rect(244, 222, 286, 240)
    text("HEATER (D) 230VAC", 265, 231, 3, "center")
    text("26 / 27", 265, 225, 2.4, "center")

    mtext("POWER NOTES (Bernard AQ SWITCH, SUG_17003 Ch.11.1):\\n"
          "1. Reversing pair -KM1 (OPEN/CCW=S1) / -KM2 (CLOSE/CW=S2). -KM2 swaps "
          "L1<->L3. SWITCH type has NO auto phase correction -> reverse by swapping "
          "two phases; phase sequence matters.\\n"
          "2. -K0 (3UG) phase-sequence/loss permissive across the incoming bus.\\n"
          "3. -Q2 (3RV) = short-circuit/backup. Motor thermal protector (thermostat, NC) "
          "is primary thermal protection and is wired into the coil control (see control "
          "sheet), per SUG_17003 7.3.\\n"
          "4. Heater (D) 230VAC kept permanently energised (anti-condensation).\\n"
          "5. Terminal numbers per TEC01-03 sheet 3.2: motor 1/2/3 +PE, "
          "thermostat 40/41, heater 26/27. NO/NC polarity within each switch "
          "group = confirm vs order sheet.",
          40, 290, 175, 3.0)

    # =====================================================================
    # ACTUATOR TERMINAL STRIP (centre)
    # =====================================================================
    text(f"{size} ACTUATOR TERMINALS (Bernard TEC01-03 sheet 3.2)", 300, 200, 4)
    rect(300, 40, 470, 192)
    # (terminal numbers, function/destination) per TEC01-03 rev06B sheet 3.2.
    # Each travel/torque is a 3-wire group: common + two contacts; NO/NC within
    # the group is order-specific -> "TBC" (see footnote).
    strip = [
        ("1 / 2 / 3  +PE", "3~ motor U/V/W (E) -> -Q2 load"),
        ("40 / 41",        "Motor thermostat (NC) -> coil string + PLC"),
        ("13 / 14 / 15",   "S2 CW close travel limit -> stops -KM2"),
        ("10 / 11 / 12",   "S1 CCW open travel limit -> stops -KM1"),
        ("7 / 8 / 9",      "Close torque -> stops -KM2 + PLC"),
        ("4 / 5 / 6",      "Open torque -> stops -KM1 + PLC"),
        ("23 / 24 / 25",   "Aux travel CLOSE - fn per order sheet (TBC)"),
        ("20 / 21 / 22",   "Aux travel OPEN - fn per order sheet (TBC)"),
        ("26 / 27",        "Heater (D) 230VAC anti-condensation"),
        ("16 / 17 / 18",   "Potentiometer (option)"),
        ("80(+) / 81(-)",  "4-20mA position xmitter (opt) 12-32VDC"),
        ("PE",             "Internal ground post"),
    ]
    yy = 184
    for term, desc in strip:
        text(term, 305, yy, 2.7)
        text(desc, 332, yy, 2.4)
        yy -= 12
    text("NO/NC within each 3-wire switch group = TBC vs order sheet", 300, 33, 2.2)

    # =====================================================================
    # CONTROL / PLC SHEET (right)
    # =====================================================================
    text("VALVE ACTUATOR  -  CONTROL & PLC I/O  (24 VDC)", 500, 600, 7)
    text(f"{pre}-Q4  5SL6216-7 (2P) + 24 VDC PSU", 500, 590, 3.4)

    line(505, 565, 505, 250)
    line(505, 565, 760, 565)
    line(500, 240, 770, 240)
    text("+24V (L+)", 507, 569, 3)
    text("0V (M)", 745, 233, 3)

    # common: +24V -> thermostat NC -> node -> two coil rungs
    line(505, 540, 540, 540)
    contact(548, 540, "nc")
    text(f"{pre}-F  THERMOSTAT 40/41 (NC)", 525, 547, 2.4)
    line(556, 540, 575, 540)
    dot(575, 540)
    line(575, 540, 575, 500)

    # OPEN rung -> -KM1
    yO = 500
    line(575, yO, 590, yO)
    contact(598, yO, "nc"); text("S1 OPEN LIMIT 10-11-12", 588, yO + 5, 2.1)
    line(606, yO, 618, yO)
    contact(626, yO, "nc"); text("OPEN TORQUE 4-5-6", 616, yO + 5, 2.1)
    line(634, yO, 646, yO)
    contact(654, yO, "nc"); text("-KM2 (ILK)", 646, yO + 5, 2.1)
    line(662, yO, 678, yO)
    rect(678, yO - 6, 694, yO + 6)
    text("-KM1", 699, yO - 1, 2.6)
    text("Q0.0 OPEN", 678, yO + 9, 2.2)
    line(694, yO, 720, yO)
    line(720, yO, 720, 240)
    dot(720, 240)

    # CLOSE rung -> -KM2
    yC = 470
    line(575, 500, 575, yC)
    line(575, yC, 590, yC)
    contact(598, yC, "nc"); text("S2 CLOSE LIMIT 13-14-15", 588, yC + 5, 2.1)
    line(606, yC, 618, yC)
    contact(626, yC, "nc"); text("CLOSE TORQUE 7-8-9", 616, yC + 5, 2.1)
    line(634, yC, 646, yC)
    contact(654, yC, "nc"); text("-KM1 (ILK)", 646, yC + 5, 2.1)
    line(662, yC, 678, yC)
    rect(678, yC - 6, 694, yC + 6)
    text("-KM2", 699, yC - 1, 2.6)
    text("Q0.1 CLOSE", 678, yC + 9, 2.2)
    line(694, yC, 730, yC)
    line(730, yC, 730, 240)
    dot(730, 240)

    text("HARDWARE STOP STRING (SUG_17003 7.3): travel-limit + torque + thermostat", 575, 512, 2.3)

    # PLC box
    rect(520, 270, 600, 440)
    mtext(f"{pre}-A1  PLC  S7-1200  CPU 1214C  (6ES7 214-1AG40)", 522, 437, 76, 2.6)
    di = [("I0.0", 420, "OPEN SIGNAL (C)"),
          ("I0.1", 405, "CLOSE SIGNAL (C)"),
          ("I0.2", 390, "OPEN TORQUE (A)"),
          ("I0.3", 375, "CLOSE TORQUE (A)"),
          ("I0.4", 360, "THERMOSTAT (NC)"),
          ("I0.5", 345, "3RV TRIP (3RV2901)"),
          ("I0.6", 330, "PHASE OK (3UG4512)"),
          ("I0.7", 315, "LOCAL / REMOTE")]
    for (ch, y, lbl) in di:
        text(ch, 523, y - 1.2, 2.5)
        line(505, y, 520, y)
        dot(505, y)
        text(lbl, 437, y + 4.5, 2.1)
        line(470, y, 478, y + 4)
    text("1M -> 0V", 525, 277, 2.3)
    line(540, 270, 540, 240)
    dot(540, 240)
    text("DO Q0.0/Q0.1 -> coil rungs above", 604, 420, 2.3)

    # HMI over PROFINET
    line(560, 270, 560, 215)
    rect(535, 195, 600, 215)
    text("ETHERNET SWITCH", 567, 204, 2.4, "center")
    line(567, 195, 567, 188)
    rect(520, 152, 600, 188)
    text("HMI   KTP700", 560, 178, 3.0, "center")
    text("OPEN / CLOSE / STOP", 560, 170, 2.2, "center")
    text("+ STATUS  (PROFINET)", 560, 162, 2.2, "center")
    text("PN", 563, 277, 2.3)

    mtext("CONTROL NOTES:\\n"
          "1. PLC outputs Q0.0 (OPEN)/Q0.1 (CLOSE) energise the -KM1/-KM2 coil rungs.\\n"
          "2. Each coil rung is broken by the actuator's own contacts - travel-limit "
          "(S1/S2), torque (open/close) and the motor thermostat - so the drive stops "
          "at end-of-travel / over-torque / over-temp even without PLC action "
          "(SUG_17003 7.3). The opposite contactor's NC aux gives the hardware interlock.\\n"
          "3. Separate signalling switches (C) report OPEN/CLOSED position to the PLC; "
          "torque, thermostat, 3RV trip and phase-OK are also PLC inputs for diagnostics.\\n"
          "4. HMI gives OPEN/CLOSE/STOP + status over PROFINET; PLC software interlock "
          "is in addition to the hardware string.\\n"
          "5. Terminal numbers per Bernard TEC01-03 sheet 3.2 (AQ SWITCH 3-ph). "
          "Only the NO/NC polarity within each switch group is TBC vs the "
          "order-specific sheet in the unit cover.",
          500, 130, 290, 3.0)

    return (f"Drew the datasheet-wired Bernard {size} valve control schematic "
            f"(power + actuator terminals + PLC/HMI, hardware-integrated limit/"
            f"torque/thermostat) at origin ({ox}, {oy}) on layer EFF-VALVE. "
            f"Terminal numbers baked in from TEC01-03 sheet 3.2; NO/NC polarity "
            f"within each switch group flagged TBC vs the order sheet.")


# ═══════════════════════════════════════════════════════════════════════════════
#  COMPLETE EFF SYSTEM SCHEMATIC  (19 sheets, NB198 corporate standard)
# ═══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def draw_eff_system(
    project_name: str = "EXTERNAL FIRE FIGHTING CONTROL SYSTEM",
    ship_name: str = "UZMAR SHIPYARD NB 198",
    drawing_no: str = "M25-IM-0001D20",
    date_str: str = "30.01.2025",
    company: str = "MARSIS MAKINE ve GEMI SAN. A.S.",
    suction_valve: str = "AQ15",
    discharge_valve: str = "AQ15",
    origin_x: float = 0.0,
    origin_y: float = 0.0,
) -> str:
    """Draw complete 19-sheet External Fire Fighting (EFF) system schematic
    matching NB198 corporate standard. Sheets: 1=Cover, 6=Block Diagram,
    7=Power Monitors, 8=Power Valves+Solenoids, 9=24VDC Distribution,
    10=PLC DQ->10K Relays, 11=PLC DI<-Actuators/Monitors,
    12=12K Relay Bank, 13=Ext Wiring Overview, 14=1X3 PORT Monitor,
    15=1X4 STBD Monitor, 16=1X5 Valve Actuators, 19=HMI+PLC Expansion.
    Components: 2x MF20EF fire monitors (each 2-axis: UP/DN + LT/RT, 0.37kW 400VAC),
    2x Bernard AQ valve actuators SWITCH type 3x400VAC (Ir=0.45A),
    HPU hydraulic clutch system, 2x S7-1200 (A1+A1.1), KTP700 HMI,
    10K relay bank Phoenix 788-312 24VDC, 12K relay bank, 3TG1010-0BB4 contactors,
    3UG4512 phase monitor, 3RV2011 motor protection. Device tags match NB198."""
    import math
    acad = _get_acad()
    doc  = acad.ActiveDocument

    for nm, builder in _SYMBOLS.items():
        if not _block_exists(doc, nm):
            try:
                builder(doc.Blocks.Add(_point(0, 0), nm))
            except Exception:
                pass
    try:
        try:
            lay = doc.Layers.Item("EFF-SYSTEM")
        except Exception:
            lay = doc.Layers.Add("EFF-SYSTEM")
            lay.color = 7
        doc.ActiveLayer = lay
    except Exception:
        pass

    ms = doc.ModelSpace
    ox, oy = float(origin_x), float(origin_y)
    REV = "00"

    def L(x1, y1, x2, y2):
        ms.AddLine(_point(x1+ox, y1+oy), _point(x2+ox, y2+oy))

    def T(txt, x, y, h=4.0, align="left"):
        _add_text(ms, str(txt), x+ox, y+oy, h, align)

    def R(x1, y1, x2, y2):
        p = ms.AddLightWeightPolyline(
            _coords([(x1+ox,y1+oy),(x2+ox,y1+oy),(x2+ox,y2+oy),(x1+ox,y2+oy)]))
        p.Closed = True

    def C(cx, cy, r):
        ms.AddCircle(_point(cx+ox, cy+oy), float(r))

    def dot(x, y, r=1.2):
        _add_dot(ms, x+ox, y+oy, r)

    def title_block(sx, sy, sheet_no, title):
        R(sx+5,  sy+5,  sx+415, sy+285)
        L(sx+5,  sy+270, sx+415, sy+270)
        L(sx+5,  sy+12,  sx+415, sy+12)
        L(sx+140,sy+270, sx+140, sy+285)
        L(sx+280,sy+270, sx+280, sy+285)
        L(sx+360,sy+270, sx+360, sy+285)
        T(company,         sx+10,  sy+278, 3.5)
        T(project_name,    sx+145, sy+279, 3.5)
        T(f"Ship: {ship_name}", sx+145, sy+273, 3)
        T(f"DWG: {drawing_no}", sx+285, sy+279, 3)
        T(f"Date: {date_str}  Rev:{REV}", sx+285, sy+273, 3)
        T(f"Sheet {sheet_no}", sx+365, sy+276, 4)
        T(title, sx+10, sy+262, 7)

    def terminal_wiring(sx, sy, tb_name, cab_terms, dev_name, dev_terms, cables):
        n = max(len(cab_terms), len(dev_terms))
        step = max(6, min(14, int(195 / max(n, 1))))
        R(sx+20, sy+50, sx+95, sy+255)
        T("CABINET",  sx+57, sy+257, 3.5, "center")
        T(tb_name,    sx+57, sy+249, 5,   "center")
        R(sx+325, sy+50, sx+410, sy+255)
        T("FIELD DEVICE", sx+367, sy+257, 3.5, "center")
        T(dev_name,       sx+367, sy+249, 4,   "center")
        for i, (tno, tlbl) in enumerate(cab_terms):
            ty = sy + 240 - i * step
            R(sx+67, ty-3, sx+83, ty+3)
            T(str(tno), sx+60, ty-1, 2.8, "right")
            T(str(tlbl)[:18], sx+86, ty-1, 2.5)
        for i, (tno, tlbl) in enumerate(dev_terms):
            ty = sy + 240 - i * step
            R(sx+325, ty-3, sx+341, ty+3)
            T(str(tno), sx+343, ty-1, 2.8)
            L(sx+83, ty, sx+325, ty)
        for j, (cno, ctype) in enumerate(cables):
            T(f"{cno}  {ctype}", sx+165, sy+248 - j*12, 3, "center")

    # ── SHEET 1: COVER ────────────────────────────────────────────────────────
    S = 0
    title_block(S, 0, "1 / 19", "COVER PAGE & SHEET INDEX")
    T("EXTERNAL FIRE FIGHTING SYSTEM", S+210, 240, 11, "center")
    T("CONTROL CABINET",               S+210, 224, 9,  "center")
    T(ship_name,                        S+210, 206, 7,  "center")
    T(company,                          S+210, 192, 5,  "center")
    T(f"Drawing No: {drawing_no}    Date: {date_str}    Rev: {REV}",
                                        S+210, 182, 4,  "center")
    R(S+15, 20, S+405, 170)
    T("SHEET INDEX", S+15, 172, 5)
    L(S+15, 162, S+405, 162)
    T("SH", S+18, 164, 3.5); T("TITLE", S+45, 164, 3.5)
    for i, (sno, desc) in enumerate([
        ("1",  "Cover Page & Sheet Index"),
        ("2",  "Remote Control Cut-Out"),
        ("3",  "Mounting Drawing"),
        ("4",  "System Overview"),
        ("5",  "Remote Control Panel Wiring"),
        ("6",  "Block Diagram"),
        ("7",  "Power Circuit – Fire Monitors (7F1, 7Q1-4, 11K1-8)"),
        ("8",  "Power Circuit – Valves & Solenoids (8Q1-2, 11K9-12, 8F1-4)"),
        ("9",  "24VDC Power Distribution (9A1 PSU, 9F1-7)"),
        ("10", "Control – PLC DQ Outputs to 10K Relay Bank"),
        ("11", "Control – PLC DI Inputs from Actuators & Monitors"),
        ("12", "Control – 12K Relay Bank (Interlocks & Alarms)"),
        ("13", "External Wiring – Cabinet Overview & Cable Schedule"),
        ("14", "External Wiring – 1X3 PORT Fire Monitor"),
        ("15", "External Wiring – 1X4 STBD Fire Monitor"),
        ("16", "External Wiring – 1X5 Suction & Discharge Valve Actuators"),
        ("17", "External Wiring – 1X2 HPU & Propulsion Signals"),
        ("18", "External Wiring – 1X1 Power In, Ship Alarm, Remote"),
        ("19", "HMI (KTP700) + PLC Expansion (A1.1, A1.2)"),
    ]):
        ry = 155 - i * 7
        L(S+15, ry+7, S+405, ry+7)
        T(sno,  S+18, ry, 3); T(desc, S+45, ry, 3)
    R(S+15, -80, S+405, 15)
    T("ELECTRICAL SPECIFICATIONS", S+18, 8, 4.5)
    for i, s in enumerate([
        f"Main Supply:        3 x 400 VAC  50 Hz  16 A",
        f"Control Supply:     24 VDC  (built-in PSU 9A1)",
        f"Suction Valve:      Bernard {suction_valve}  SWITCH type  3x400VAC  Ir=0.45A",
        f"Discharge Valve:    Bernard {discharge_valve}  SWITCH type  3x400VAC  Ir=0.45A",
        f"Fire Monitors:      2x MF20EF  (PORT + STBD)  0.37kW / axis  400VAC",
        f"Motor Protection:   3RV2011-1AA10  (valves: 0.45A) + 3RV2011 (monitors: 1.1A)",
        f"Contactors:         3TG1010-0BB4  24VDC coil  (11K1-11K12)",
        f"Control Relays:     Phoenix Contact 788-312  24VDC  (10K1-9, 12K1-7)",
        f"PLC:                Siemens S7-1200 CPU 1214C + SM1223 expansions (A1.1, A1.2)",
        f"HMI:                Siemens KTP700  7\" PROFINET",
    ]):
        T(s, S+20, -2 - i*8, 3.2)

    # ── SHEET 6: BLOCK DIAGRAM ────────────────────────────────────────────────
    S = 2300
    title_block(S, 0, "6 / 19", "SYSTEM BLOCK DIAGRAM")
    R(S+155, 85, S+265, 245)
    T("MAIN CONTROL", S+210, 238, 4.5, "center")
    T("CABINET",      S+210, 228, 4.5, "center")
    for j, ln in enumerate(["S7-1200 A1+A1.1+A1.2","KTP700 HMI (PROFINET)",
                             "400V→24VDC PSU (9A1)","10K/12K relay banks",
                             "11K1-12 power contactors","1X1-1X6 terminals"]):
        T(ln, S+160, 215-j*10, 3)
    def _blk(lbl, lines, x1, y1, x2, y2, wx, wy, dx, dy, cno, ctype):
        R(S+x1, y1, S+x2, y2)
        T(lbl, S+(x1+x2)//2, y2-10, 4, "center")
        for k, ln in enumerate(lines):
            T(ln, S+(x1+x2)//2, y2-20-k*9, 3, "center")
        L(S+wx, wy, S+wx+dx, wy+dy)
        T(f"{cno} {ctype}", S+wx+(3 if dx>0 else -55), wy+2, 2.8)
    _blk("PORT MONITOR",["MF20EF","UP/DN 0.37kW","LT/RT 0.37kW"],
         15,175,135,230, 135,202, 20,0, "1W3-1/2","7x+16x1.5")
    _blk("STBD MONITOR",["MF20EF","UP/DN 0.37kW","LT/RT 0.37kW"],
         285,175,405,230, 265,202,-20,0,"1W4-1/2","7x+16x1.5")
    _blk("SUCTION VALVE",[f"Bernard {suction_valve}","SWITCH 3x400V"],
         15,105,135,155, 135,130, 20,0,"1W5-1/2","7x1.5+4x2x0.75")
    _blk("DISCHARGE VLV",[f"Bernard {discharge_valve}","SWITCH 3x400V"],
         285,105,405,155, 265,130,-20,0,"1W5-3/4","7x1.5+4x2x0.75")
    _blk("HPU / CLUTCH",["400V pump motor","oil press alarm"],
         90,30,200,75, 155,95, 0,-10,"1W2","3x1.5+4x2x0.75")
    _blk("REMOTE PANEL",["Bridge joysticks","monitor buttons"],
         220,30,330,75, 265,85,-10,0,"1W6","7x2x0.75+CAT8")
    T("Ship Supply 3x400V 16A in → 1X1", S+160, 250, 3.5)
    T("1W1-6 → Ship Alarm  |  1W1-3/4/5 → Propulsion  |  1W6-4 CAT8 → PROFINET",
      S+10, -15, 3.2)

    # ── SHEET 7: POWER – FIRE MONITORS ───────────────────────────────────────
    S = 2760
    title_block(S, 0, "7 / 19", "POWER CIRCUIT – FIRE MONITORS  400V 50Hz")
    L(S+15, 258, S+410, 258)
    T("400V 50Hz  from 1X1", S+15, 260, 3.5)
    for col, lbl in ((S+60,"L1"),(S+75,"L2"),(S+90,"L3")):
        L(col, 268, col, 258); T(lbl, col-2, 269, 3.5); dot(col, 258)
    R(S+48, 243, S+102, 258)
    T("7F1  3P  16A", S+75, 250, 3.5, "center")
    for col in (S+60, S+75, S+90):
        L(col, 243, col, 237)
    L(S+15, 237, S+410, 237)

    def mon_branch(bx, qno, ka, kb, tbname, t_start, ax_a, ax_b):
        cx = S+bx; c1=cx-5; c2=cx; c3=cx+5
        for c in (c1,c2,c3): L(c, 237, c, 217); dot(c, 237)
        R(cx-12, 195, cx+12, 217)
        T(f"7{qno}", cx, 211, 3.5, "center")
        T("3RV Ir=1.1A", cx, 201, 2.8, "center")
        for c in (c1,c2,c3): L(c, 195, c, 178)
        R(cx-12, 158, cx+12, 178)
        T(ka, cx, 171, 3.5, "center"); T("3TG1010", cx, 162, 2.5, "center")
        T(ax_a, cx, 158, 3, "center")
        for c in (c1,c2,c3): L(c, 158, c, 140)
        R(cx-12, 120, cx+12, 140)
        T(kb, cx, 133, 3.5, "center"); T("3TG1010", cx, 124, 2.5, "center")
        T(ax_b, cx, 120, 3, "center")
        for c in (c1,c2,c3): L(c, 120, c, 108)
        L(c1,108,c1,100); L(c3,108,c3,100); L(c2,108,c2,100)
        L(c1,100,c3,94); L(c3,100,c1,94); L(c2,100,c2,94)
        R(cx-12, 82, cx+12, 94)
        T(tbname, cx, 90, 2.8, "center")
        T(f":{t_start}-{t_start+2}", cx, 84, 2.8, "center")
        for c in (c1,c2,c3): L(c, 95, c, 94); L(c, 82, c, 72)
        C(cx, 60, 10); T("M", cx-2, 57, 4); T("3~", cx-2, 50, 3)
        T(f"0.37kW {ax_a}/{ax_b}", cx, 43, 2.8, "center")

    mon_branch(60,  "Q1","11K1","11K2","1X3",1,"UP","DN")
    mon_branch(150, "Q2","11K3","11K4","1X3",4,"LT","RT")
    mon_branch(255, "Q3","11K5","11K6","1X4",1,"UP","DN")
    mon_branch(345, "Q4","11K7","11K8","1X4",4,"LT","RT")
    L(S+15,35,S+195,35); T("PORT FIRE MONITOR (MF20EF)", S+25, 29, 4)
    L(S+215,35,S+410,35); T("STBD FIRE MONITOR (MF20EF)", S+225, 29, 4)
    T("NOTE: 11K1-K8 coils 24VDC (3TG1010-0BB4). Hardware interlock: opposite NC aux in coil rung.",
      S+15, 18, 2.8)

    # ── SHEET 8: POWER – VALVES + SOLENOIDS ──────────────────────────────────
    S = 3220
    title_block(S, 0, "8 / 19", "POWER CIRCUIT – VALVES & SOLENOIDS  400V 50Hz")
    L(S+15, 258, S+410, 258)
    T("400V 50Hz  (cross-ref sheet 7)", S+15, 260, 3.5)
    for col, lbl in ((S+55,"L1"),(S+65,"L2"),(S+75,"L3")):
        L(col, 268, col, 258); T(lbl, col-2, 269, 3.5); dot(col, 258)
    R(S+90, 248, S+200, 260)
    T("8A1  3UG4512  PHASE MONITOR", S+145, 254, 3, "center")

    def valve_branch(bx, qno, kopen, kclose, t_motor, t_htr, vlbl, vsz):
        cx=S+bx; c1=cx-4.3; c2=cx; c3=cx+4.3
        for c in (c1,c2,c3): L(c, 258, c, 235); dot(c, 258)
        R(cx-12, 213, cx+12, 235)
        T(f"8{qno}", cx, 228, 3.5, "center")
        T("3RV2011-1AA10", cx, 219, 2.5, "center")
        T("Ir=0.45A set", cx, 213, 2.8, "center")
        for c in (c1,c2,c3): L(c, 213, c, 196)
        R(cx-12, 176, cx+12, 196)
        T(kopen, cx, 188, 3.5, "center"); T("3TG1010", cx, 180, 2.5, "center")
        T("OPEN", cx, 176, 3, "center")
        for c in (c1,c2,c3): L(c, 176, c, 159)
        R(cx-12, 139, cx+12, 159)
        T(kclose, cx, 151, 3.5, "center"); T("3TG1010", cx, 143, 2.5, "center")
        T("CLOSE", cx, 139, 3, "center")
        for c in (c1,c2,c3): L(c, 139, c, 126)
        R(cx-12, 112, cx+12, 124)
        T("1X5", cx, 121, 2.8, "center"); T(t_motor, cx, 114, 2.8, "center")
        for c in (c1,c2,c3): L(c, 126, c, 124); L(c, 112, c, 100)
        C(cx, 88, 10); T("M", cx-2, 85, 4); T("3~", cx-2, 78, 3)
        T(f"Bernard {vsz}", cx, 70, 3, "center"); T(vlbl, cx, 63, 3, "center")
        L(cx+12, 224, cx+35, 224); L(cx+35, 224, cx+35, 50)
        R(cx+28, 40, cx+55, 50)
        T("HTR 400V", cx+41, 46, 2.5, "center"); T(t_htr, cx+41, 40, 2.5, "center")

    valve_branch(65,  "Q1","11K9", "11K10","1,2,3",  ":5,6",   "SUCTION",   suction_valve)
    valve_branch(175, "Q2","11K11","11K12","14,15,16",":18,19", "DISCHARGE", discharge_valve)
    for fx, fno, flbl in [(S+245,"8F1","PORT FOG/JET"),(S+285,"8F2","PORT WTR/FOAM"),
                           (S+325,"8F3","STBD FOG/JET"),(S+365,"8F4","STBD WTR/FOAM")]:
        L(fx, 268, fx, 258); dot(fx, 258)
        R(fx-8, 247, fx+8, 258); T(fno, fx, 253, 2.8, "center"); T("2A", fx, 248, 2.8, "center")
        L(fx, 247, fx, 232); R(fx-10, 218, fx+10, 232)
        T(flbl, fx, 226, 2.5, "center"); T("solenoid", fx, 220, 2.5, "center")
        T("→1X2", fx, 214, 2.5, "center")
    T("NOTE: 11K9-12 hardware interlock (travel limits+thermostat) in coil rung – see sheet 11.",
      S+15, 18, 2.8)

    # ── SHEET 9: 24VDC DISTRIBUTION ──────────────────────────────────────────
    S = 3680
    title_block(S, 0, "9 / 19", "24VDC POWER DISTRIBUTION")
    R(S+145, 210, S+275, 260)
    T("9A1  24VDC POWER SUPPLY", S+210, 253, 3.5, "center")
    T("IN: 3x400VAC  OUT: 24VDC 16A", S+210, 244, 3, "center")
    L(S+210, 268, S+210, 260); T("3x400V from 1X1", S+213, 263, 3)
    L(S+145, 232, S+15, 232); L(S+275, 232, S+410, 232); T("24VDC L+", S+17, 234, 3.5)
    L(S+145, 222, S+15, 222); L(S+275, 222, S+410, 222); T("0V M",     S+17, 224, 3.5)
    R(S+15, 207, S+80, 222)
    T("9F1  16A MCB", S+47, 215, 3, "center"); T("MAIN CTRL BUS", S+47, 208, 3, "center")
    L(S+47, 222, S+47, 232); L(S+47, 207, S+47, 195)
    L(S+15, 195, S+410, 195); L(S+15, 55, S+410, 55)
    T("24VDC L+  (distributed)", S+17, 197, 3.5); T("0V M", S+17, 57, 3.5)
    for fx, fno, fa, flbl in [
        (S+40,  "9F2","4A","KTP700 HMI"),   (S+90,  "9F3","4A","HYD PUMP"),
        (S+140, "9F4","4A","PORT MON CTRL"), (S+195, "9F5","4A","STBD MON CTRL"),
        (S+250, "9F6","6A","SOLENOID OUT"),  (S+310, "9F7","4A","REMOTE+ALARM"),
        (S+365, "9F8","4A","PLC CPU SUPPLY")]:
        L(fx, 195, fx, 183); R(fx-10, 170, fx+10, 183)
        T(fno, fx, 179, 2.8, "center"); T(fa, fx, 171, 2.8, "center")
        L(fx, 170, fx, 145); R(fx-12, 120, fx+12, 145)
        T(flbl, fx, 133, 3, "center"); L(fx, 120, fx, 55)
    T("Fuse type: Phoenix Contact 2002-xxxx 24VDC miniature.", S+15, 43, 3)

    # ── SHEET 10: PLC DQ → 10K RELAY BANK ───────────────────────────────────
    S = 4140
    title_block(S, 0, "10 / 19", "CONTROL – PLC DQ OUTPUTS TO 10K RELAY BANK")
    L(S+15, 255, S+410, 255); T("24VDC L+  from 9F6/9F4/9F5", S+15, 257, 3.5)
    L(S+15, 45,  S+410, 45);  T("0V M", S+15, 47, 3.5)
    R(S+15, 175, S+95, 248)
    T("S7-1200 A1", S+55, 241, 4, "center"); T("DQ module", S+55, 232, 3.5, "center")
    for i, (dq, desc, kno, action) in enumerate([
        ("Q0.0","PORT FOG ON",   "10K1","→ PORT FOG solenoid via 8F1"),
        ("Q0.1","PORT JET ON",   "10K2","→ PORT JET solenoid via 8F1"),
        ("Q0.2","PORT WATER ON", "10K3","→ PORT WATER solenoid via 8F2"),
        ("Q0.3","PORT FOAM ON",  "10K4","→ PORT FOAM solenoid via 8F2"),
        ("Q0.4","STBD FOG ON",   "10K5","→ STBD FOG solenoid via 8F3"),
        ("Q0.5","STBD JET ON",   "10K6","→ STBD JET solenoid via 8F3"),
        ("Q0.6","STBD WATER ON", "10K7","→ STBD WATER solenoid via 8F4"),
        ("Q0.7","STBD FOAM ON",  "10K8","→ STBD FOAM solenoid via 8F4"),
        ("Q1.0","HYD CLUTCH ON", "10K9","→ HPU clutch coil 24VDC via 9F3"),
    ]):
        rx = S+115+i*33
        L(rx, 255, rx, 200); T(dq, rx-3, 227, 2.8)
        R(rx-8, 185, rx+8, 200); T(kno, rx, 195, 3, "center"); T("788-312", rx, 187, 2.2, "center")
        L(rx, 185, rx, 45); T(action, rx-10, 170, 2.5)
    T("10K TYPE: Phoenix Contact 788-312  24VDC coil  SPDT  250VAC 6A", S+15, 35, 3.2)
    T("10K NO contacts → solenoid load circuits (via fuses 8F1-8F4).", S+15, 27, 3)

    # ── SHEET 11: PLC DI INPUTS ───────────────────────────────────────────────
    S = 4600
    title_block(S, 0, "11 / 19", "CONTROL – PLC DI INPUTS  (ACTUATORS & MONITORS)")
    L(S+15, 255, S+410, 255); T("24VDC L+", S+15, 257, 3.5)
    L(S+15, 45,  S+410, 45);  T("0V M", S+15, 47, 3.5)
    R(S+320, 90, S+410, 255)
    T("S7-1200 A1", S+365, 249, 4, "center"); T("DI module", S+365, 240, 3.5, "center")
    T("PLC DI ASSIGNMENTS (A1 module):", S+15, 250, 4.5)
    for i, (di, desc, src) in enumerate([
        ("I0.0","SUCTION OPENED",   "1X5:11 NO travel limit"),
        ("I0.1","SUCTION CLOSED",   "1X5:14 NC travel limit"),
        ("I0.2","DISCHARGE OPENED", "1X5:24 NO travel limit"),
        ("I0.3","DISCHARGE CLOSED", "1X5:27 NC travel limit"),
        ("I0.4","PORT MON LS1",     "1X3:11 limit switch 1"),
        ("I0.5","PORT MON LS2",     "1X3:14 limit switch 2"),
        ("I0.6","STBD MON LS1",     "1X4:11 limit switch 1"),
        ("I0.7","STBD MON LS2",     "1X4:14 limit switch 2"),
        ("I1.0","PHASE OK",         "8A1 output NC (opens on fault)"),
        ("I1.1","HPU OVERLOAD",     "1X2:9 NC contact 3RV aux"),
        ("I1.2","HPU OIL PRESS",    "1X2:10 NC pressure switch"),
        ("I1.3","HYD CLUTCH ON",    "1X2:11 NO feedback contact"),
        ("I1.4","16S1 SELECTOR",    "OFF / SEMI / AUTO switch"),
        ("I1.5","17S1 SELECTOR",    "PORT / STBD select switch"),
    ]):
        ry = 240 - i*13
        L(S+15, ry, S+35, ry); R(S+35, ry-3, S+48, ry+3); L(S+48, ry, S+320, ry)
        T(di, S+18, ry+2, 3); T(desc, S+52, ry+2, 3.2); T(src, S+52, ry-5, 2.8)
    T("VALVE COIL RUNG – hardware stop (Bernard SUG_17003 §7.3):", S+15, 57, 3.5)
    T("L+ → 12K6 NO → 1X5:10-11 travel OPEN NC → 1X5:40-41 thermostat NC → 11K10 NC → 11K9 coil → M",
      S+15, 49, 3)
    T("L+ → 12K6 NC → 1X5:13-14 travel CLOSE NC → 1X5:40-41 thermostat NC → 11K9 NC → 11K10 coil → M",
      S+15, 41, 3)
    T("Same structure for DISCHARGE valve (12K7, 11K11, 11K12).", S+15, 33, 3)

    # ── SHEET 12: 12K RELAY BANK ──────────────────────────────────────────────
    S = 5060
    title_block(S, 0, "12 / 19", "CONTROL – 12K RELAY BANK  (INTERLOCKS & ALARMS)")
    L(S+15, 255, S+410, 255); T("24VDC L+", S+15, 257, 3.5)
    L(S+15, 45,  S+410, 45);  T("0V M", S+15, 47, 3.5)
    for i, (kno, lbl, io_ref, action) in enumerate([
        ("12K1","FIFI MODE\nREADY",    "Q1.1","NO→1X1 prop sig A"),
        ("12K2","CLUTCH\nENGAGED",     "Q1.2","NO→1X1 prop sig B"),
        ("12K3","M/E START\nINTERLOCK","I2.0","NC in M/E control circuit"),
        ("12K4","COMMON\nALARM",       "I2.1","NO→ship alarm 1W1-6"),
        ("12K5","BUZZER",              "I2.2","NO→panel buzzer B1"),
        ("12K6","SUCTION\nVLV CMD",    "Q2.0","NO→11K9(OPEN) NC→11K10(CLOSE)"),
        ("12K7","DISCHARGE\nVLV CMD",  "Q2.1","NO→11K11(OPEN) NC→11K12(CLOSE)"),
    ]):
        rx = S+30+i*55
        L(rx, 255, rx, 200); T(io_ref, rx-5, 230, 2.8)
        R(rx-10, 183, rx+10, 200)
        T(kno, rx, 194, 3.5, "center"); T("788-312", rx, 185, 2.2, "center")
        L(rx, 183, rx, 45)
        for j, ln in enumerate(lbl.split("\n")): T(ln, rx-8, 174-j*9, 3)
        T(action, rx-12, 160, 2.3)
    T("12K3 NC: prevents M/E start while FIFI active. 12K4 NO: common alarm to ship.",
      S+15, 35, 3)
    T("12K6/12K7 pulse cmds from PLC; hardware travel limits in contactor coil rung.",
      S+15, 27, 3)

    # ── SHEET 13: EXTERNAL WIRING OVERVIEW ───────────────────────────────────
    S = 5520
    title_block(S, 0, "13 / 19", "EXTERNAL WIRING – CABINET OVERVIEW & CABLE SCHEDULE")
    R(S+155, 100, S+265, 245)
    T("MAIN CONTROL", S+210, 238, 4.5, "center"); T("CABINET", S+210, 228, 4.5, "center")
    for tb, desc in [("1X1","400V IN+alarm"),("1X2","HPU+propulsion"),
                     ("1X3","PORT monitor"), ("1X4","STBD monitor"),
                     ("1X5","Valve actuators"),("1X6","Remote panel")]:
        T(f"{tb}: {desc}", S+160, 215-["1X1","1X2","1X3","1X4","1X5","1X6"].index(tb)*12, 3)
    T("CABLE SCHEDULE", S+15, 92, 5)
    R(S+15, 10, S+415, 92)
    c_hdr = [S+15,S+55,S+100,S+160,S+220,S+290,S+365,S+415]
    for cx, h in zip(c_hdr, ["CABLE","FROM","TO","TYPE","CORES","FUNCTION","SH"]):
        T(h, cx+2, 84, 3.2); L(cx, 10, cx, 92)
    for i, row in enumerate([
        ("1W1-1","1X1","SUPPLY","3x2.5mm²","3","400VAC power in","7,8"),
        ("1W1-2","1X1","PSU","2x2.5mm²","2","24V PSU feed","9"),
        ("1W1-3","1X2","PROPULSION","2x2x0.75","4","M/E interlock","12"),
        ("1W1-4","1X2","PROPULSION","4x2x0.75","8","Speed/mode sigs","12"),
        ("1W1-5","1X2","PROPULSION","2x2x0.75","4","FIFI mode out","12"),
        ("1W1-6","1X1","SHIP ALARM","2x2x0.75","4","Common alarm","12"),
        ("1W2-1","1X2","HPU","3x1.5mm²","3","HPU 400V pump","8"),
        ("1W2-2","1X2","HPU","2x1.5mm²","2","Clutch 24VDC","10"),
        ("1W2-3","1X2","HPU","4x2x0.75","8","HPU status sigs","11"),
        ("1W3-1","1X3","PORT MON","7x1.5mm²","7","UP/DN motor","7"),
        ("1W3-2","1X3","PORT MON","16x1.5mm²","16","LT/RT motor","7"),
        ("1W3-3","1X3","PORT MON","2x2x0.75","4","Limit sw sigs","11"),
        ("1W4-1","1X4","STBD MON","7x1.5mm²","7","UP/DN motor","7"),
        ("1W4-2","1X4","STBD MON","16x1.5mm²","16","LT/RT motor","7"),
        ("1W4-3","1X4","STBD MON","2x2x0.75","4","Limit sw sigs","11"),
        ("1W5-1","1X5","SUCT VLV","7x1.5mm²","7","Motor+heater SUCT","8,16"),
        ("1W5-2","1X5","SUCT VLV","4x2x0.75","8","Limit sw SUCT","11,16"),
        ("1W5-3","1X5","DISCH VLV","7x1.5mm²","7","Motor+heater DISCH","8,16"),
        ("1W5-4","1X5","DISCH VLV","4x2x0.75","8","Limit sw DISCH","11,16"),
        ("1W6-1","1X6","REMOTE","7x2x0.75","14","Monitor cmds","10"),
        ("1W6-2","1X6","REMOTE","7x2x0.75","14","Monitor feedback","11"),
        ("1W6-3","1X6","REMOTE","2x2.5mm²","2","Remote 24VDC supply","9"),
        ("1W6-4","1X6","HMI/PLC","CAT7/CAT8","8","PROFINET Ethernet","19"),
    ]):
        ry = 78 - i*3.0
        L(S+15, ry+3, S+415, ry+3)
        for cx, val in zip(c_hdr, row): T(val, cx+2, ry, 2.5)

    # ── SHEET 14: 1X3 PORT MONITOR WIRING ────────────────────────────────────
    S = 5980
    title_block(S, 0, "14 / 19", "EXTERNAL WIRING – 1X3  PORT FIRE MONITOR")
    terminal_wiring(S, 0, "1X3",
        [(1,"L1 UP/DN motor"),(2,"L2 UP/DN motor"),(3,"L3 UP/DN motor"),
         (4,"L1 LT/RT motor"),(5,"L2 LT/RT motor"),(6,"L3 LT/RT motor"),
         (7,"PE UP/DN"),(8,"PE LT/RT"),(9,"PE aux"),(10,"24VDC +"),
         (11,"LS1 OPEN a-cont"),(12,"LS1 common"),(13,"LS2 common"),
         (14,"LS2 CLOSE a-cont"),(15,"LS3"),(16,"LS4"),(17,"0V return"),
         (18,"Maint button +"),(19,"Maint button −"),
         (20,"FOG sol +"),(21,"FOG sol −"),(22,"JET sol +")],
        "PORT MON MF20EF",
        [(1,"U UP/DN"),(2,"V UP/DN"),(3,"W UP/DN"),
         (4,"U LT/RT"),(5,"V LT/RT"),(6,"W LT/RT"),
         (7,"PE"),(8,"PE"),(9,"PE"),(10,"24V+"),
         (11,"LS1-a"),(12,"LS1-C"),(13,"LS2-C"),(14,"LS2-a"),
         (15,"LS3"),(16,"LS4"),(17,"0V"),
         (18,"MB+"),(19,"MB−"),(20,"FOG+"),(21,"FOG−"),(22,"JET+")],
        [("1W3-1","7x1.5mm²"),("1W3-2","16x1.5mm²"),("1W3-3","2x2x0.75mm² shld")]
    )
    T("1W3-1 (7x1.5mm²): UP/DN motor L1/L2/L3/PE + spare | 1W3-2 (16x1.5mm²): LT/RT + aux",
      S+15, -10, 3)
    T("1W3-3 (2x2x0.75mm² shielded): limit switch signals BU/BK colour-coded pairs",
      S+15, -18, 3)

    # ── SHEET 15: 1X4 STBD MONITOR WIRING ───────────────────────────────────
    S = 6440
    title_block(S, 0, "15 / 19", "EXTERNAL WIRING – 1X4  STBD FIRE MONITOR")
    terminal_wiring(S, 0, "1X4",
        [(1,"L1 UP/DN motor"),(2,"L2 UP/DN motor"),(3,"L3 UP/DN motor"),
         (4,"L1 LT/RT motor"),(5,"L2 LT/RT motor"),(6,"L3 LT/RT motor"),
         (7,"PE UP/DN"),(8,"PE LT/RT"),(9,"PE aux"),(10,"24VDC +"),
         (11,"LS1 OPEN a-cont"),(12,"LS1 common"),(13,"LS2 common"),
         (14,"LS2 CLOSE a-cont"),(15,"LS3"),(16,"LS4"),(17,"0V return"),
         (18,"Maint button +"),(19,"Maint button −"),
         (20,"FOG sol +"),(21,"FOG sol −"),(22,"JET sol +")],
        "STBD MON MF20EF",
        [(1,"U UP/DN"),(2,"V UP/DN"),(3,"W UP/DN"),
         (4,"U LT/RT"),(5,"V LT/RT"),(6,"W LT/RT"),
         (7,"PE"),(8,"PE"),(9,"PE"),(10,"24V+"),
         (11,"LS1-a"),(12,"LS1-C"),(13,"LS2-C"),(14,"LS2-a"),
         (15,"LS3"),(16,"LS4"),(17,"0V"),
         (18,"MB+"),(19,"MB−"),(20,"FOG+"),(21,"FOG−"),(22,"JET+")],
        [("1W4-1","7x1.5mm²"),("1W4-2","16x1.5mm²"),("1W4-3","2x2x0.75mm² shld")]
    )
    T("1W4-1 (7x1.5mm²): UP/DN motor | 1W4-2 (16x1.5mm²): LT/RT + aux",  S+15, -10, 3)
    T("1W4-3 (2x2x0.75mm² shielded): limit switch signals", S+15, -18, 3)

    # ── SHEET 16: 1X5 SUCTION + DISCHARGE VALVE WIRING ──────────────────────
    S = 6900
    title_block(S, 0, "16 / 19", "EXTERNAL WIRING – 1X5  SUCTION & DISCHARGE VALVE ACTUATORS")
    R(S+15, 50, S+90, 250)
    T("CABINET", S+52, 252, 3.5, "center"); T("1X5", S+52, 244, 5, "center")
    x5_terms = [
        (1,"SUCT L1","1W5-1"),   (2,"SUCT L2","1W5-1"),
        (3,"SUCT L3","1W5-1"),   (4,"SUCT PE","1W5-1"),
        (5,"SUCT HTR L","1W5-1"),(6,"SUCT HTR N","1W5-1"),
        (7,"SUCT T-OPN C","1W5-2"),(8,"SUCT T-OPN NC","1W5-2"),
        (9,"SUCT T-OPN NO","1W5-2"),(10,"SUCT T-CLS C","1W5-2"),
        (11,"SUCT T-CLS NC","1W5-2"),(12,"SUCT THERM +","1W5-2"),
        (13,"SUCT THERM -","1W5-2"),
        (14,"DISCH L1","1W5-3"),(15,"DISCH L2","1W5-3"),
        (16,"DISCH L3","1W5-3"),(17,"DISCH PE","1W5-3"),
        (18,"DISCH HTR L","1W5-3"),(19,"DISCH HTR N","1W5-3"),
        (20,"DISCH T-OPN C","1W5-4"),(21,"DISCH T-OPN NC","1W5-4"),
        (22,"DISCH T-OPN NO","1W5-4"),(23,"DISCH T-CLS C","1W5-4"),
        (24,"DISCH T-CLS NC","1W5-4"),(25,"DISCH T-CLS NO","1W5-4"),
        (26,"DISCH THERM +","1W5-4"),
    ]
    step16 = 7
    for i, (tno, tlbl, cbl) in enumerate(x5_terms):
        ty = 236 - i*step16
        R(S+62, ty-3, S+78, ty+3)
        T(str(tno), S+55, ty-1, 3, "right"); T(tlbl[:18], S+80, ty-1, 2.5)
        T(cbl, S+130, ty-1, 2.5); L(S+15, ty, S+62, ty)
    R(S+290, 140, S+415, 253)
    T(f"SUCTION VALVE", S+352, 248, 3.5, "center")
    T(f"Bernard {suction_valve}", S+352, 239, 4.5, "center")
    T("SWITCH  3x400VAC  (TEC01-03 sh.3.2)", S+352, 230, 2.8, "center")
    for j, (tno, tlbl) in enumerate([(1,"Motor L1"),(2,"Motor L2"),(3,"Motor L3"),
            ("PE","PE"),(26,"Heater L"),(27,"Heater N"),
            (10,"Travel OPEN C"),(11,"Travel OPEN NC"),(12,"Travel OPEN NO"),
            (13,"Travel CLOSE C"),(14,"Travel CLOSE NC"),(15,"Travel CLOSE NO"),
            (40,"Thermostat NC +"),(41,"Thermostat NC -")]):
        ty2 = 218 - j*5.5
        R(S+290, ty2-2.5, S+303, ty2+2.5)
        T(str(tno), S+305, ty2-2, 2.5); T(tlbl, S+316, ty2-2, 2.3)
    R(S+290, 30, S+415, 135)
    T(f"DISCHARGE VALVE", S+352, 130, 3.5, "center")
    T(f"Bernard {discharge_valve}", S+352, 121, 4.5, "center")
    T("SWITCH  3x400VAC", S+352, 112, 2.8, "center")
    for j, (tno, tlbl) in enumerate([(1,"Motor L1"),(2,"Motor L2"),(3,"Motor L3"),
            ("PE","PE"),(26,"Heater L"),(27,"Heater N"),
            (10,"Travel OPEN C"),(11,"Travel OPEN NC"),(12,"Travel OPEN NO"),
            (13,"Travel CLOSE C"),(14,"Travel CLOSE NC"),(15,"Travel CLOSE NO"),
            (40,"Thermostat NC +"),(41,"Thermostat NC -")]):
        ty2 = 108 - j*5.5
        R(S+290, ty2-2.5, S+303, ty2+2.5)
        T(str(tno), S+305, ty2-2, 2.5); T(tlbl, S+316, ty2-2, 2.3)
    for i, (tno, tlbl, _) in enumerate(x5_terms):
        ty = 236 - i*step16
        dev_y = (218 - i*5.5) if i <= 12 else (108 - (i-13)*5.5)
        L(S+78, ty, S+290, dev_y)
    T("1W5-1 (7x1.5mm²): SUCT motor L1/L2/L3/PE + heater → AQ terms 1,2,3,PE,26,27",  S+15, 22, 3)
    T("1W5-2 (4x2x0.75mm²): SUCT travel limits 10-15 + thermostat 40-41 → AQ terms",   S+15, 14, 3)
    T("1W5-3 (7x1.5mm²): DISCH motor L1/L2/L3/PE + heater → AQ terms 1,2,3,PE,26,27", S+15,  6, 3)
    T("1W5-4 (4x2x0.75mm²): DISCH travel limits 10-15 + thermostat 40-41",              S+15, -2, 3)

    # ── SHEET 19: HMI + PLC EXPANSION ────────────────────────────────────────
    S = 8280
    title_block(S, 0, "19 / 19", "HMI (KTP700) + PLC EXPANSION  (A1.1 / A1.2)")
    R(S+15, 155, S+145, 258)
    T("KTP700 HMI", S+80, 252, 5, "center")
    T("6AV2 123-2GB03-0AX0", S+80, 242, 3, "center")
    T("7\" TFT  PROFINET RJ45", S+80, 233, 3.5, "center")
    T("Supply: 24VDC (9F2)", S+80, 223, 3.5, "center")
    T("FUNCTIONS:", S+20, 212, 3.5)
    for j, fn in enumerate(["PORT monitor: UP/DN/LT/RT","STBD monitor: UP/DN/LT/RT",
                              "Valve OPEN/CLOSE/STATUS","FOG/JET/WATER/FOAM select",
                              "HPU start/stop + clutch ON","FIFI mode enable/disable",
                              "Alarm acknowledge"]):
        T(f"  {fn}", S+22, 203-j*9, 3)
    R(S+160, 190, S+245, 258)
    T("S7-1200 CPU (A1)", S+202, 252, 4, "center")
    T("6ES7 214-1AG40-0XB0",S+202, 242, 2.8, "center")
    T("DI 14  I0.0-I1.5",   S+202, 232, 3.2, "center")
    T("DQ 10  Q0.0-Q1.1",   S+202, 222, 3.2, "center")
    T("PROFINET port (see sh.10,11)", S+202, 210, 2.8, "center")
    R(S+255, 190, S+325, 258)
    T("A1.1  SM1223", S+290, 252, 4, "center"); T("8DI / 8DQ", S+290, 242, 3.5, "center")
    T("DI: I2.0-I3.7", S+290, 232, 3.2, "center"); T("DQ: Q2.0-Q3.7", S+290, 222, 3.2, "center")
    T("11K contactor", S+290, 212, 3, "center"); T("coil outputs", S+290, 204, 3, "center")
    R(S+335, 190, S+405, 258)
    T("A1.2  SM1223", S+370, 252, 4, "center"); T("8DI / 8DQ", S+370, 242, 3.5, "center")
    T("DI: I4.0-I4.7", S+370, 232, 3.2, "center"); T("DQ: Q4.0-Q4.7", S+370, 222, 3.2, "center")
    T("12K relay", S+370, 212, 3, "center"); T("coil outputs", S+370, 204, 3, "center")
    for px in (S+80, S+202, S+290, S+370): L(px, 190, px, 181)
    L(S+80, 181, S+370, 181)
    R(S+78, 170, S+375, 181)
    T("PROFINET / Ethernet  (1W6-4  CAT7/CAT8)", S+225, 174, 3.5, "center")
    L(S+80, 258, S+80, 268); L(S+202, 258, S+202, 268); L(S+80, 268, S+202, 268)
    T("24VDC: 9F2 (KTP700)  9F8 (PLC via 1X1)", S+80, 270, 3)
    T("TOTAL I/O: DI=38  DQ=34 across A1+A1.1+A1.2", S+15, 158, 4.5)
    R(S+15, 80, S+405, 155)
    for cx, h in zip([S+15,S+95,S+140,S+250],[" MODULE","I/O","COUNT","FUNCTION SUMMARY"]):
        T(h, cx+2, 147, 3.5); L(cx, 80, cx, 155)
    L(S+15, 143, S+405, 143)
    for i, row in enumerate([
        ("A1 CPU 1214C","DI","I0.0-I1.5  14","Valve limits, monitor LS, phase, HPU, selectors"),
        ("A1 CPU 1214C","DQ","Q0.0-Q1.1  10","10K1-9 solenoid/clutch relays, FIFI mode signals"),
        ("A1.1 SM1223", "DI","I2.0-I3.7  16","Propulsion, M/E speed, clutch, ship systems"),
        ("A1.1 SM1223", "DQ","Q2.0-Q3.7  16","11K1-12 contactor coils (monitors + valves)"),
        ("A1.2 SM1223", "DI","I4.0-I4.7   8","Remote panel, maintenance buttons"),
        ("A1.2 SM1223", "DQ","Q4.0-Q4.7   8","12K3-7 interlock/alarm/buzzer relay coils"),
    ]):
        ry = 135 - i*10; L(S+15, ry+10, S+405, ry+10)
        for cx, val in zip([S+15,S+95,S+140,S+250], row): T(val, cx+2, ry+1, 3)

    # ── Create AutoCAD layouts ────────────────────────────────────────────────
    sheet_positions = [
        ("EFF-S01-Cover",        0,    ),
        ("EFF-S06-Block Diag",   2300, ),
        ("EFF-S07-Pwr Monitors", 2760, ),
        ("EFF-S08-Pwr Valves",   3220, ),
        ("EFF-S09-24VDC",        3680, ),
        ("EFF-S10-PLC DQ",       4140, ),
        ("EFF-S11-PLC DI",       4600, ),
        ("EFF-S12-12K Relays",   5060, ),
        ("EFF-S13-Ext Overview", 5520, ),
        ("EFF-S14-1X3 Port",     5980, ),
        ("EFF-S15-1X4 Stbd",     6440, ),
        ("EFF-S16-1X5 Valves",   6900, ),
        ("EFF-S19-HMI+PLC",      8280, ),
    ]
    created = []
    for lname, lsx in sheet_positions:
        try:
            try:
                lo = doc.Layouts.Item(lname)
            except Exception:
                lo = doc.Layouts.Add(lname)
            try:
                lo.PlotPaperSize = (420.0, 297.0); lo.PaperUnits = 1
            except Exception:
                pass
            try:
                import time
                doc.ActiveLayout = lo
                time.sleep(0.25)
                doc.SendCommand("MSPACE\n")
                time.sleep(0.25)
                x1 = lsx + ox - 5;  y1 = oy - 5
                x2 = lsx + ox + 425; y2 = oy + 290
                doc.SendCommand(f"ZOOM\nW\n{x1},{y1}\n{x2},{y2}\n")
                time.sleep(0.25)
                doc.SendCommand("PSPACE\n")
                time.sleep(0.15)
            except Exception:
                pass
            created.append(lname)
        except Exception as e:
            created.append(f"{lname}(ERR:{e})")
    try:
        doc.Regen(1)
    except Exception:
        pass

    return (
        f"Drew complete EFF system schematic – 13 sheets for '{project_name}' / '{ship_name}'. "
        f"Valve actuators: {suction_valve} (suction), {discharge_valve} (discharge), "
        f"3RV2011-1AA10 Ir=0.45A. "
        f"Fire monitors: 2x MF20EF 0.37kW 400VAC 4-axis. "
        f"Contactors: 3TG1010-0BB4 24VDC coil 11K1-11K12. "
        f"Control relays: Phoenix Contact 788-312 24VDC 10K1-9, 12K1-7. "
        f"PLCs: S7-1200 A1+A1.1+A1.2 (38DI/34DQ). HMI: KTP700 PROFINET. "
        f"Layouts created: {'; '.join(created)}. "
        f"All sheets in model space at origin ({origin_x},{origin_y}), "
        f"460-unit horizontal spacing. Sheet 16 maps 1X5 terminals to "
        f"Bernard AQ terminal numbers (TEC01-03 sheet 3.2)."
    )


# ═══════════════════════════════════════════════════════════════════════════════
#  SINGLE VALVE CONTROL SCHEMATIC  –  örnek style, A1 template-based
# ═══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def draw_single_valve(
    aq_model: str = "AQ25",
    valve_label: str = "SEA WATER VALVE",
    tag_prefix: str = "1",
    main_mcb: str = "5SL6316-7",
    ctrl_mcb: str = "5SL6216-7",
    plc_model: str = "6ES7 214-1AG40-0XB0",
    project_name: str = "VALVE CONTROL SYSTEM",
    ship_name: str = "",
    drawing_no: str = "M00-IM-0001",
    date_str: str = "",
    company: str = "",
    output_path: str = "",
) -> str:
    """Draw single valve control schematic in NB198/ornek style on ISO A1 sheet.
    Uses ornek valf cizimi.dwg as template for CAD blocks (BENEK, BOBIN, KON3P,
    CB_TM, 3P FUSE, NA, NK, KLESIG, OK, wire, role aciklama).
    One compact A1 sheet: power section (MCB->phase monitor->motor prot->contactors),
    control section (PLC S7-1200 DI/DO, coil rungs with NC interlocks),
    terminal/wiring section (1X1->AQ field, cable labels).
    Motor protection auto-selected from Bernard AQ datasheet (3x400VAC 50Hz).
    Supports AQ5 through AQ1000."""
    import os, shutil, time, datetime
    import win32com.client
    import pythoncom

    # Bernard AQ lookup: (In_A, Istart_A, kW, torque_Nm, time_s, rv_part, rv_range, rv_set)
    # Source: Bernard Controls TEC B60ver_PRG_F+E, 3x400VAC 50Hz
    _AQ = {
        "AQ5":    (0.16, 0.43, 0.03,    50,  16, "3RV2011-0FA10", "0.35-0.50", "0.40"),
        "AQ10":   (0.16, 0.43, 0.03,   100,  25, "3RV2011-0FA10", "0.35-0.50", "0.40"),
        "AQ15":   (0.16, 0.43, 0.03,   150,  30, "3RV2011-0FA10", "0.35-0.50", "0.40"),
        "AQ25":   (0.22, 0.43, 0.04,   250,  30, "3RV2011-0GA10", "0.45-0.63", "0.45"),
        "AQ30":   (0.22, 0.43, 0.04,   300,  35, "3RV2011-0GA10", "0.45-0.63", "0.45"),
        "AQ50":   (0.43, 0.75, 0.07,   500,  35, "3RV2011-1BA10", "0.55-0.80", "0.60"),
        "AQ80":   (0.43, 0.75, 0.07,   800,  55, "3RV2011-1BA10", "0.55-0.80", "0.60"),
        "AQ150":  (0.97, 3.80, 0.40,  1500,  40, "3RV2011-1EA10", "1.10-1.60", "1.10"),
        "AQ280":  (0.97, 3.80, 0.40,  2800, 100, "3RV2011-1EA10", "1.10-1.60", "1.10"),
        "AQ430":  (2.70,14.00, 0.90,  4300,  70, "3RV2021-4AA10", "2.80-4.00", "3.00"),
        "AQ610":  (2.70,14.00, 0.90,  6100, 100, "3RV2021-4AA10", "2.80-4.00", "3.00"),
        "AQ830":  (1.90, 6.70, 0.80,  8300, 115, "3RV2011-1GA10", "1.80-2.50", "2.00"),
        "AQ1000": (1.60, 5.00, 0.50, 10400, 135, "3RV2011-1FA10", "1.40-2.00", "1.60"),
    }

    aq_key = aq_model.strip().upper().replace(" ", "")
    if aq_key not in _AQ:
        return ("ERROR: unknown AQ model '" + aq_model + "'. "
                "Supported: " + ", ".join(sorted(_AQ.keys())))
    In, Istart, kW, torque, t_s, rv_part, rv_range, rv_set = _AQ[aq_key]

    if not date_str:
        date_str = datetime.date.today().strftime("%d.%m.%Y")
    P = str(tag_prefix)

    # ── Template setup ───────────────────────────────────────────────────────
    # The template file already has all corporate blocks (BENEK, BOBIN, etc.)
    import glob as _glob
    matches = _glob.glob(r"c:\Users\dogao\Downloads\*rnek*valf*.dwg")
    TEMPLATE = matches[0] if matches else r"c:\Users\dogao\Downloads\ornek valf cizimi.dwg"
    if not os.path.exists(TEMPLATE):
        return "ERROR: Template not found. Put 'ornek valf cizimi.dwg' in Downloads folder."

    if not output_path:
        out_dir = os.path.dirname(TEMPLATE)
        safe = valve_label.replace(" ", "_").replace("/", "-")[:20]
        output_path = os.path.join(out_dir, f"Valve_{P}_{aq_key}_{safe}.dwg")

    try:
        acad = win32com.client.GetActiveObject("ZwCAD.Application")
    except Exception as e:
        return "ERROR: ZWCAD not running – " + str(e)

    # ── Copy template to ASCII-safe temp path ────────────────────────────
    # Avoids Turkish-char encoding issues in ZWCAD command line
    import tempfile as _tmplib
    _tmp_tmpl = os.path.join(_tmplib.gettempdir(), "zwcad_valve_template.dwg")
    shutil.copy(TEMPLATE, _tmp_tmpl)
    time.sleep(0.1)

    # ── Open / locate the output document ────────────────────────────────
    _needed_set = {"BENEK","OK","NA","NK","KON3P","KLESIG",
                   "3P FUSE","CB_TM","BOBIN","role aciklama","wire"}

    def _blk_names(d):
        names = set()
        for _bi in range(d.Blocks.Count):
            try: names.add(d.Blocks.Item(_bi).Name)
            except Exception: pass
        return names

    # Priority 1: output_path already open — reuse it
    doc = None
    for _i in range(acad.Documents.Count):
        _d = acad.Documents.Item(_i)
        try:
            if os.path.normcase(_d.FullName) == os.path.normcase(output_path):
                doc = _d
                break
        except Exception:
            pass

    # Priority 2: another open valve drawing that already has the blocks
    if doc is None:
        for _i in range(acad.Documents.Count):
            _d = acad.Documents.Item(_i)
            _dn = _d.Name.lower()
            if "rnek" in _dn or "nb198" in _dn:
                continue  # keep örnek and system drawing untouched
            if _needed_set & _blk_names(_d):
                doc = _d
                break

    # Priority 3: try to open a fresh copy of the template
    if doc is None:
        shutil.copy(_tmp_tmpl, output_path)
        time.sleep(0.2)
        try:
            doc = acad.Documents.Open(output_path)
            time.sleep(0.8)
        except Exception:
            # ZWCAD at document limit — close a non-essential doc first
            for _i in range(acad.Documents.Count - 1, -1, -1):
                _d = acad.Documents.Item(_i)
                _dn = _d.Name.lower()
                if "rnek" not in _dn and "nb198" not in _dn:
                    try:
                        _d.Close(False)
                        time.sleep(0.4)
                    except Exception:
                        pass
                    break
            try:
                doc = acad.Documents.Open(output_path)
                time.sleep(0.8)
            except Exception:
                return "ERROR: Cannot open output document in ZWCAD (doc limit)."

    if doc is None:
        return "ERROR: Could not obtain a ZWCAD document."

    ms = doc.ModelSpace

    # Erase all existing model-space entities (block definitions are preserved)
    for i in range(ms.Count - 1, -1, -1):
        try:
            ms.Item(i).Delete()
        except Exception:
            pass

    # ── Ensure corporate block definitions exist ───────────────────────────
    missing_blocks = [b for b in sorted(_needed_set)
                      if b not in _blk_names(doc)]

    if missing_blocks:
        # Strategy A: find örnek in open docs, copy block defs via COM
        tmpl_doc = None
        for _i in range(acad.Documents.Count):
            _d = acad.Documents.Item(_i)
            if "rnek" in _d.Name.lower() or "valf" in _d.Name.lower():
                tmpl_doc = _d
                break
        if tmpl_doc is not None:
            src_blk_objs = []
            _mb_set = set(missing_blocks)
            for _bi in range(tmpl_doc.Blocks.Count):
                try:
                    _blk = tmpl_doc.Blocks.Item(_bi)
                    if _blk.Name in _mb_set and not _blk.Name.startswith("*"):
                        src_blk_objs.append(_blk)
                except Exception:
                    pass
            if src_blk_objs:
                try:
                    _objs_var = win32com.client.VARIANT(
                        pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, src_blk_objs)
                    tmpl_doc.CopyObjects(_objs_var, doc.Blocks.ObjectID)
                    time.sleep(0.3)
                    missing_blocks = []  # mark done
                except Exception:
                    pass  # fall to Strategy B

        if missing_blocks:
            # Strategy B: -INSERT blockname=ascii_temp_path via SendCommand
            doc.Activate()
            time.sleep(0.2)
            for _bname in missing_blocks:
                try:
                    doc.SendCommand(f'-INSERT\n"{_bname}"={_tmp_tmpl}\n')
                    time.sleep(0.3)
                    doc.SendCommand("\x1b\n")
                    time.sleep(0.15)
                except Exception:
                    pass
            time.sleep(0.3)

    # ── Drawing helpers ──────────────────────────────────────────────────────
    def _pt(x, y):
        return win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(x), float(y), 0.0])

    def L(x1, y1, x2, y2):
        ms.AddLine(_pt(x1, y1), _pt(x2, y2))

    def C(cx, cy, r):
        ms.AddCircle(_pt(cx, cy), float(r))

    def A(cx, cy, r, s=0, e=360):
        import math
        ms.AddArc(_pt(cx, cy), float(r), math.radians(float(s)), math.radians(float(e)))

    def T(s, x, y, h=0.25):
        ms.AddText(str(s), _pt(x, y), float(h))

    def BLK(name, x, y, sx=1.0, sy=1.0, rot=0.0):
        try:
            return ms.InsertBlock(_pt(x, y), str(name), float(sx), float(sy), 1.0, float(rot))
        except Exception:
            return None

    # Shorthand block helpers using örnek block names
    def dot(x, y):         BLK("BENEK", x, y)            # junction dot
    def nc_contact(x, y):  BLK("NA", x, y)               # NC contact symbol
    def relay_nc(x, y):    BLK("NK", x, y)               # relay NC contact
    def coil(x, y):        BLK("BOBIN", x, y)            # coil
    def desc_box(x, y):    BLK("role aciklama", x, y)    # relay description box
    def terminal(x, y):    BLK("OK", x, y)               # connection terminal
    def term_blk(x, y):    BLK("KLESIG", x, y)           # terminal block symbol
    def cable_lbl(x, y):   BLK("wire", x, y)             # cable label
    def mcb_3p(x, y):      BLK("3P FUSE", x, y)         # 3-pole MCB
    def motor_prot(x, y):  BLK("CB_TM", x, y)           # motor protection CB
    def contactor(x, y):   BLK("KON3P", x, y)           # 3-pole contactor

    # ── COORDINATE CONSTANTS  (örnek scale: 60:1 on A1) ────────────────────
    # Power section — exact values from örnek valf çizimi.dwg block/wire analysis
    X_MCB = 195.74; Y_MCB = 111.09  # "3P FUSE" block insertion (centre pole = L2)
    X_MP  = 198.77; Y_MP  = 111.09  # "CB_TM" block insertion (centre pole = L2)

    # Phase relay
    X_RELAY = 203.0;  Y_RELAY = 111.1
    X_RELAY_L = X_RELAY - 0.5;  X_RELAY_R = X_RELAY + 0.9
    Y_RELAY_T = Y_RELAY + 0.45;  Y_RELAY_B = Y_RELAY - 0.45

    X_K1  = 198.77; X_K2  = 200.72; Y_K = 107.31  # contactor block insertions
    X_MOT = 198.77; Y_MOT = 103.29  # motor circle centre, r=0.55

    # Coil columns — vertical rungs (matching örnek's vertical layout)
    X_K1C = 217.23; X_K2C = 218.43  # K1 OPEN / K2 CLOSE coil column X positions
    Y_RUNG_TOP = 104.79  # K1/K2 rung top — PLC DQ output connection level
    Y_NK   = 102.84  # electrical interlock NK contact Y
    Y_COIL = 100.0   # BOBIN block centre Y

    Y_BOT = 99.0;  Y_TOP = 113.5

    # ═══════════════════════════════════════════════════════════════════════
    # POWER SECTION  (X ≈ 195–208)
    # Layout matches örnek exactly: 3 staggered supply buses (L1/L2/L3 at
    # different Y), CB_TM poles at X=198.37/198.77/199.17, reversing pair
    # K1/K2 with phase swap for motor direction.
    # ═══════════════════════════════════════════════════════════════════════

    # Pole X positions — 3P FUSE poles (insertion at X_MCB, centre = L2)
    X_F_L1 = X_MCB - 0.40   # 195.34
    X_F_L2 = X_MCB           # 195.74
    X_F_L3 = X_MCB + 0.40   # 196.14

    # CB_TM + K1 contactor poles (insertion at X_MP = X_K1, centre = L2)
    X_L1 = X_MP - 0.40      # 198.37
    X_L2 = X_MP              # 198.77
    X_L3 = X_MP + 0.40      # 199.17

    # K2 contactor poles (insertion at X_K2, centre = L2)
    X_K2_L1 = X_K2 - 0.40   # 200.32
    X_K2_L2 = X_K2            # 200.72
    X_K2_L3 = X_K2 + 0.40   # 201.12

    # Staggered supply bus Y (one horizontal bus per phase, from örnek wire data)
    Y_BUS_L1 = 112.74
    Y_BUS_L2 = 112.39
    Y_BUS_L3 = 112.04

    # Key Y levels from örnek wire analysis
    Y_FUSE_BOT  = 110.49   # 3P FUSE load-side terminal
    Y_CB_BOT_L1 = 109.54   # CB_TM load-side L1 (block geometry)
    Y_CB_BOT_L2 = 109.69   # CB_TM load-side L2
    Y_CB_BOT_L3 = 109.54   # CB_TM load-side L3
    Y_CROSS_L1  = 108.06   # K1–K2 cross-bus L1 (distributes CB_TM output to both contactors)
    Y_CROSS_L2  = 108.46   # K1–K2 cross-bus L2
    Y_CROSS_L3  = 108.86   # K1–K2 cross-bus L3
    Y_K_BOT     = 106.71   # contactor load-side output terminal
    Y_MOT_L1    = 105.20   # motor junction L1 (phase-swap cross point)
    Y_MOT_L2    = 105.60   # motor junction L2
    Y_MOT_L3    = 106.00   # motor junction L3
    Y_MOT_ENTRY = 104.49   # motor terminal entry Y
    Y_HTR_KLESIG = 107.24  # heater KLESIG terminal strip Y

    # ── Three staggered supply buses (horizontal, L1/L2/L3 at different Y) ──
    T("400V 50Hz", 199.81, 112.89, 0.22)   # at supply bus right end (örnek position)
    T("400V 50Hz", 195.22, 103.89, 0.18)   # at cable entry bottom (örnek position)
    T("SUPPLY",    195.32, 103.59, 0.16)   # cable entry label (örnek position)
    # Supply cable entry labels (örnek positions at Y=104.54)
    T("R", 195.44, 104.54, 0.18)
    T("S", 195.84, 104.54, 0.18)
    T("T", 196.24, 104.54, 0.18)
    # Cable X1 references (örnek positions)
    T("1X1", 194.76, 104.36, 0.14)
    T("1X1", 207.89, 103.47, 0.14)   # control section entry (örnek second instance)
    L(X_F_L1, Y_BUS_L1, 205.1, Y_BUS_L1)
    L(X_F_L2, Y_BUS_L2, 204.7, Y_BUS_L2)
    L(X_F_L3, Y_BUS_L3, 204.3, Y_BUS_L3)

    # ── F1 Main 3P FUSE — incoming line fuse ─────────────────────────────
    mcb_3p(X_MCB, Y_MCB)
    T(P + "F1", X_MCB - 0.8, Y_MCB + 0.55, 0.22)
    T(main_mcb, X_MCB - 0.8, Y_MCB + 0.30, 0.18)
    T("16 A", 196.42, 110.68, 0.20)   # örnek position (right of fuse block)
    T("16 A", 207.92, 110.76, 0.20)   # örnek second instance (control entry)
    # Fuse → supply bus (each pole at its own bus Y)
    L(X_F_L1, Y_BUS_L1, X_F_L1, Y_MCB)
    L(X_F_L2, Y_BUS_L2, X_F_L2, Y_MCB)
    L(X_F_L3, Y_BUS_L3, X_F_L3, 111.01)  # L3 terminal at 111.01 (not Y_MCB 111.09)
    # Incoming supply from cable entry (below fuse load side → Y_MOT_ENTRY)
    for xf in (X_F_L1, X_F_L2, X_F_L3):
        L(xf, Y_FUSE_BOT, xf, Y_MOT_ENTRY)

    # ── Phase relay A1 box ────────────────────────────────────────────────
    # Phase relay A1 — labels only (no box in örnek)
    T(P + "A1", X_RELAY_L + 0.05, Y_RELAY_T + 0.12, 0.20)
    T(P + "A1", 203.23, 105.61, 0.20)   # second instance at cable entry (örnek)
    T("PHASE CONTROL", 203.75, 104.89, 0.16)
    # L1/L2/L3 supply connections at right end of buses (vertical lines)
    for xr, ybus in ((205.1, Y_BUS_L1), (204.7, Y_BUS_L2), (204.3, Y_BUS_L3)):
        L(xr, 105.88, xr, ybus)
    # Phase relay input labels at right end of buses (örnek exact positions)
    T("L1", 204.19, 105.6, 0.18)
    T("L2", 204.57, 105.6, 0.18)
    T("L3", 204.97, 105.6, 0.18)

    # ── Q1 Motor protection (CB_TM) ───────────────────────────────────────
    # Poles at X_L1/L2/L3; each pole taps the supply bus with a BENEK junction
    motor_prot(X_MP, Y_MP)
    T(P + "Q1", X_MP + 0.8, Y_MP + 0.55, 0.22)
    T(rv_part, X_MP + 0.8, Y_MP + 0.30, 0.18)
    _rv_r = rv_range.replace('.', ','); _rv_s = rv_set.replace('.', ',')
    T(f"Ir={_rv_r}A", X_MP + 0.86, Y_MP - 0.95, 0.20)
    T(f"set={_rv_s}A", X_MP + 0.86, Y_MP - 1.25, 0.20)
    # CB_TM line-side to supply bus (BENEK at junction)
    L(X_L1, Y_MCB, X_L1, Y_BUS_L1);  dot(X_L1, Y_BUS_L1)
    L(X_L2, Y_MCB, X_L2, Y_BUS_L2);  dot(X_L2, Y_BUS_L2)
    L(X_L3, Y_MCB, X_L3, Y_BUS_L3);  dot(X_L3, Y_BUS_L3)
    # CB_TM load-side to K1 input (continuous wire through cross-bus BENEK)
    L(X_L1, Y_CB_BOT_L1, X_L1, Y_K)
    L(X_L2, Y_CB_BOT_L2, X_L2, 107.24)  # L2 K1 input at 107.24 (block geometry)
    L(X_L3, Y_CB_BOT_L3, X_L3, Y_K)

    # ── K1–K2 cross-bus (distributes CB_TM output to both contactors) ────
    L(X_L1, Y_CROSS_L1, X_K2_L1, Y_CROSS_L1);  dot(X_L1, Y_CROSS_L1)
    L(X_L2, Y_CROSS_L2, X_K2_L2, Y_CROSS_L2);  dot(X_L2, Y_CROSS_L2)
    L(X_L3, Y_CROSS_L3, X_K2_L3, Y_CROSS_L3);  dot(X_L3, Y_CROSS_L3)

    # ── K1 OPEN contactor ────────────────────────────────────────────────
    contactor(X_K1, Y_K)
    T(P + "K1", X_K1 - 0.8, Y_K + 0.30, 0.22)

    # ── K2 CLOSE contactor ───────────────────────────────────────────────
    contactor(X_K2, Y_K)
    T(P + "K2", X_K2 + 0.5, Y_K + 0.30, 0.22)

    # K2 input wires from cross-bus to contactor top
    L(X_K2_L1, Y_K, X_K2_L1, Y_CROSS_L1)
    L(X_K2_L2, Y_K, X_K2_L2, Y_CROSS_L2)
    L(X_K2_L3, Y_K, X_K2_L3, Y_CROSS_L3)

    # ── Contactor outputs + phase swap for reversing ──────────────────────
    # K2 outputs go down to phase-swap junction Y
    L(X_K2_L1, Y_K_BOT, X_K2_L1, Y_MOT_L3)  # K2 L1 → motor L3 junction
    L(X_K2_L2, Y_K_BOT, X_K2_L2, Y_MOT_L2)  # K2 L2 → motor L2 junction
    L(X_K2_L3, Y_K_BOT, X_K2_L3, Y_MOT_L1)  # K2 L3 → motor L1 junction
    # Horizontal phase-swap cross-wires between K1 and K2 outputs
    L(X_K2_L1, Y_MOT_L3, X_L3, Y_MOT_L3)   # K2 L1 ↔ K1 L3
    L(X_K2_L2, Y_MOT_L2, X_L2, Y_MOT_L2)   # K2 L2 ↔ K1 L2
    L(X_K2_L3, Y_MOT_L1, X_L1, Y_MOT_L1)   # K2 L3 ↔ K1 L1
    # K1 output wires: from K1 load-side down to motor cable connector top (Y=104.49)
    L(X_L1, Y_K_BOT, X_L1, Y_MOT_ENTRY)
    L(X_L2, Y_K_BOT, X_L2, Y_MOT_ENTRY)
    L(X_L3, Y_K_BOT, X_L3, Y_MOT_ENTRY)
    # BENEK at motor junctions (where K2 cross-wire meets K1 output)
    dot(X_L1, Y_MOT_L1)
    dot(X_L2, Y_MOT_L2)
    dot(X_L3, Y_MOT_L3)
    # Motor entry stubs (gap Y=104.49→104.39 is cable connector plate) + diagonals
    L(X_L1, 104.39, X_L1, 103.96)               # L1 stub below connector
    L(X_L2, 103.66, X_L2, 104.39)               # L2 stub
    L(X_L3, 104.39, X_L3, 103.96)               # L3 stub
    L(198.53, 103.61, X_L1, 103.96)             # L1 diagonal cable entry
    L(199.01, 103.61, X_L3, 103.96)             # L3 diagonal cable entry
    # Motor terminal numbers at connector plate (örnek exact positions)
    T("14", 198.47, 104.54, 0.14)
    T("15", 198.87, 104.54, 0.14)
    T("16", 199.27, 104.54, 0.14)
    T("3",  198.52, 102.86, 0.14)
    # Motor cable reference
    T("1X5", 197.8, 104.36, 0.14)

    # ── Motor symbol ──────────────────────────────────────────────────────
    C(X_MOT, Y_MOT, 0.55)
    T("M", X_MOT - 0.14, Y_MOT - 0.10, 0.30)
    T("3~", X_MOT - 0.15, Y_MOT - 0.38, 0.18)
    T(valve_label.upper(), X_MOT - 0.55, Y_MOT - 0.85, 0.20)
    T("OPEN CLOSE", X_MOT - 0.68, Y_MOT - 1.48, 0.16)
    # Motor terminal connection arcs (from örnek ET4 entities)
    A(198.85, 102.96, 0.075)
    A(199.0,  102.96, 0.075)

    # ── Heater — taps directly from L2 and L3 supply buses ───────────────
    # BENEK at bus tap points; KLESIG terminals at Y=107.24
    X_HTR_L1 = 201.59   # heater wire 1 taps from L2 bus
    X_HTR_L2 = 201.99   # heater wire 2 taps from L3 bus
    dot(X_HTR_L1, Y_BUS_L2)
    dot(X_HTR_L2, Y_BUS_L3)
    L(X_HTR_L1, Y_BUS_L2, X_HTR_L1, Y_MOT_ENTRY)  # heater L1: bus to connector top
    L(X_HTR_L2, Y_BUS_L3, X_HTR_L2, Y_MOT_ENTRY)  # heater L2: bus to connector top
    term_blk(X_HTR_L1, Y_HTR_KLESIG)
    term_blk(X_HTR_L2, Y_HTR_KLESIG)
    # Lower heater cable entry section (matching örnek; gap Y=104.49→104.39 = connector plate)
    L(X_HTR_L1, 103.63, X_HTR_L1, 104.39)       # HTR L1 upper stub below connector
    L(X_HTR_L2, 103.63, X_HTR_L2, 104.39)       # HTR L2 upper stub
    L(X_HTR_L1, 102.06, X_HTR_L1, 103.53)       # HTR L1 lower section
    L(X_HTR_L2, 102.81, X_HTR_L2, 103.53)       # HTR L2 middle section
    L(X_HTR_L2, 102.06, X_HTR_L2, 102.21)       # HTR L2 bottom stub
    L(X_HTR_L1, 102.06, X_HTR_L2, 102.06)       # horizontal at cable entry bottom
    T(P + "F3", X_HTR_L1 - 0.1, Y_HTR_KLESIG + 0.55, 0.18)
    T(P + "F4", X_HTR_L2 - 0.1, Y_HTR_KLESIG + 0.55, 0.18)   # second heater fuse
    T("2A",     X_HTR_L1 + 0.1, Y_HTR_KLESIG + 0.15, 0.14)   # heater fuse current rating
    T("HEATER", 201.42, 101.31, 0.16)
    T("400VAC", 201.45, 101.61, 0.16)
    # Heater terminal numbers at connector plate (örnek exact positions)
    T("17", 201.69, 104.49, 0.14)
    T("18", 202.09, 104.49, 0.14)
    T("26", 201.7,  103.47, 0.14)   # heater cable terminal numbers
    T("27", 202.1,  103.47, 0.14)
    # Heater cable reference
    T("1X5", 201.02, 104.2, 0.14)

    # ═══════════════════════════════════════════════════════════════════════
    # CONTROL SECTION  (X ≈ 208–227)  — exact örnek coordinates
    # ═══════════════════════════════════════════════════════════════════════

    # Bus Y levels from örnek
    Y_LP   = 112.75   # L+ incoming (unprotected) bus
    Y_PROT = 112.36   # Protected 24V bus (after protection contacts)
    Y_0V   = 99.3     # External 0V M bus
    Y_0VI  = 99.69    # Internal 0V bus (status + coil common)

    # ── L+ and protected buses ────────────────────────────────────────────
    L(208.46, Y_LP,   235.42, Y_LP)     # full L+ bus to OK terminal
    L(208.81, Y_LP,   234.25, Y_LP)     # duplicate L+ bus segment (örnek has both)
    L(208.81, Y_PROT, 234.25, Y_PROT)   # protected bus
    L(209.98, Y_0V,   218.43, Y_0V)     # external 0V bus
    T("24VDC", 221.85, 112.96, 0.22)

    # ── Supply terminal labels (SUPPLY / 24VDC at X≈208) ─────────────────
    T("SUPPLY", 208.25, 102.7, 0.16)
    T("24VDC",  208.32, 103.0, 0.16)
    T("+", 208.56, 103.65, 0.18)
    T("-", 208.91, 103.65, 0.18)

    # ── Protection contact 1: thermostat NA at (208.46, 111.31) ──────────
    L(208.46, Y_LP, 208.46, 111.31)       # L+ bus down to thermostat top
    BLK('NA', 208.46, 111.31)
    T("1", 208.51, 111.16, 0.11)
    T("2", 208.51, 110.76, 0.11)
    L(208.44, 110.91, 208.79, 110.91)     # bridge to phase relay
    L(208.46, 110.78, 208.46, 103.6)      # thermostat bottom wire down to supply

    # Thermostat device symbol (zig-zag lines at Y≈111)
    for _s, _e in [
        ((208.15, 111.0),  (208.21, 111.0)),
        ((208.19, 111.06), (208.14, 111.03)),
        ((208.2,  111.05), (208.15, 111.0)),
        ((208.23, 111.08), (208.24, 111.06)),
        ((208.26, 111.0),  (208.26, 111.03)),
        ((208.26, 111.03), (208.25, 111.05)),
        ((208.31, 111.11), (208.23, 111.08)),
        ((208.32, 111.09), (208.31, 111.11)),
        ((208.36, 111.1),  (208.32, 111.09)),
        ((208.37, 111.08), (208.26, 111.03)),
    ]:
        L(_s[0], _s[1], _e[0], _e[1])

    # ── Protection contact 2: phase relay NA at (208.81, 111.31) ─────────
    L(208.81, 111.31, 208.81, Y_PROT)     # phase relay top to protected bus
    BLK('NA', 208.81, 111.31)
    T("3", 208.86, 111.16, 0.11)
    T("4", 208.86, 110.76, 0.11)
    L(208.81, 110.71, 208.81, 103.6)      # phase relay bottom wire down

    # Phase relay device symbol
    for _s, _e in [
        ((208.55, 111.0),  (208.61, 111.0)),
        ((208.58, 111.08), (208.59, 111.06)),
        ((208.59, 111.06), (208.54, 111.03)),
        ((208.6,  111.05), (208.55, 111.0)),
        ((208.61, 111.0),  (208.61, 111.03)),
        ((208.61, 111.03), (208.6,  111.05)),
        ((208.66, 111.11), (208.58, 111.08)),
        ((208.67, 111.09), (208.66, 111.11)),
        ((208.71, 111.1),  (208.67, 111.09)),
        ((208.72, 111.08), (208.61, 111.03)),
    ]:
        L(_s[0], _s[1], _e[0], _e[1])

    # ── Control section boundary lines ────────────────────────────────────
    L(209.18, Y_0VI, 214.83, Y_0VI)       # horizontal boundary at Y=99.69
    L(209.18, Y_LP,  209.18, Y_0VI)       # left vertical wall
    L(209.98, Y_PROT, 209.98, Y_0V)       # vertical at X=209.98

    # ── PLC DI column verticals + short stubs ─────────────────────────────
    L(210.33, 105.53, 210.33, 106.53)
    L(210.63, 106.93, 210.63, Y_LP)       # DI column 1
    L(211.23, 106.93, 211.23, Y_PROT)     # DI column 2
    L(211.83, 106.93, 211.83, Y_PROT)     # DI column 3

    # PLC DI pin stubs (4 pins at X=211.04/211.33/211.61/211.89)
    for _xp in (211.04, 211.33, 211.61, 211.89):
        L(_xp, 105.13, _xp, 104.85)       # top stub
        L(_xp, 104.62, _xp, 103.32)       # bottom stub
        A(_xp, 104.75, 0.1)               # pin connection arc (örnek ET4)
    L(211.89, 103.32, 211.04, 103.32)     # horizontal connector
    L(211.47, 103.32, 211.47, 102.12)     # down to OK terminal
    terminal(211.47, 102.12)              # 0V OK terminal

    # PLC labels
    T("L+",  210.51, 106.69, 0.18)
    T("M",   211.14, 106.69, 0.18)
    T("M",   211.74, 106.69, 0.18)
    T("A1",  210.42, 105.79, 0.18)
    T("ETHERNET CONNECTION", 210.66, 105.29, 0.16)
    T("RJ45", 212.13, 104.71, 0.16)
    T("FROM TOUCH SCREEN", 210.29, 101.65, 0.16)
    T("GW",  211.01, 103.61, 0.16)
    T("OW",  211.29, 103.61, 0.16)
    T("G",   211.58, 103.61, 0.16)
    T("O",   211.86, 103.61, 0.16)
    T("CAT7", 211.56, 102.85, 0.14)
    T("CAT8", 211.56, 103.09, 0.14)
    T("O  :", 212.4, 103.82, 0.12)
    T("G  :", 212.4, 103.97, 0.12)
    T("OW :", 212.4, 104.13, 0.12)
    T("GW :", 212.4, 104.28, 0.12)
    T("ORANGE",       212.77, 103.82, 0.12)
    T("GREEN",        212.77, 103.97, 0.12)
    T("ORANGE WHITE", 212.77, 104.13, 0.12)
    T("GREEN WHITE",  212.77, 104.28, 0.12)

    # ── OPENED status rung (X=213.63) — 6 segments matching örnek ────────
    L(213.63, 111.44, 213.63, Y_LP)
    L(213.63, 110.64, 213.63, 111.34)
    BLK('NA', 213.63, 110.14)
    L(213.63, 110.14, 213.63, 110.54)
    L(213.63, 109.14, 213.63, 109.54)
    L(213.63, 108.35, 213.63, 109.04)
    L(213.63, 108.24, 213.63, 106.93)
    T("0.0",  213.48, 106.65, 0.14)
    T(valve_label.upper(), 213.1, 108.69, 0.14)
    T("OPENED", 213.38, 109.54, 0.14)
    T("1X5",  213.06, 108.22, 0.12)
    T("1X5",  213.06, 111.31, 0.12)
    T("24",   213.78, 108.22, 0.12)
    T("22",   213.78, 108.99, 0.12)
    T("20",   213.78, 110.49, 0.12)
    T("23",   213.78, 111.31, 0.12)

    # ── CLOSED status rung (X=214.83) — 7 segments matching örnek ────────
    L(214.83, 111.44, 214.83, Y_LP)
    L(214.83, 110.64, 214.83, 111.34)
    BLK('NA', 214.83, 110.14)
    L(214.83, 110.14, 214.83, 110.54)
    L(214.83, 109.14, 214.83, 109.54)
    L(214.83, 108.35, 214.83, 109.04)
    L(214.83, 108.24, 214.83, 106.93)
    L(214.83, 105.13, 214.83, Y_0VI)       # continues down to internal 0V bus
    T("0.1",  214.68, 106.65, 0.14)
    T("CLOSED", 214.58, 109.55, 0.14)
    T("1X5",  214.26, 108.22, 0.12)
    T("1X5",  214.26, 111.31, 0.12)
    T("26",   214.98, 108.22, 0.12)
    T("25",   214.98, 108.99, 0.12)
    T("23",   214.98, 110.49, 0.12)
    T("25",   214.98, 111.31, 0.12)
    T("1L",   214.71, 105.25, 0.14)

    # ── OL_AUX rung (X=216.03) — 2 segments matching örnek ──────────────
    L(216.03, 110.19, 216.03, Y_LP)
    BLK('NA', 216.03, 110.19)
    L(216.03, 109.59, 216.03, 106.93)
    T("0.2",  215.88, 106.65, 0.14)
    T("8A1",  215.46, 109.84, 0.14)
    T("11",   216.08, 110.09, 0.14)
    T("14",   216.08, 109.64, 0.14)

    # ── DQ output bus (örnek: X=214.83 to X=226.83 at Y=104.79) ─────────
    L(226.83, Y_RUNG_TOP, 214.83, Y_RUNG_TOP)
    L(226.83, Y_RUNG_TOP, 226.83, 105.13)

    # ── K1 OPEN coil rung (X=217.23) — 8 segments matching örnek ────────
    L(217.23, 104.46, 217.23, 105.13)
    L(217.23, 104.36, 217.23, 103.14)
    BLK('NK', 217.23, Y_NK)
    L(217.23, 103.04, 217.23, Y_NK)
    L(217.23, 102.24, 217.23, 102.04)
    L(217.23, 101.94, 217.23, 100.76)
    BLK('BOBIN', 217.23, Y_COIL)
    L(217.23, 100.66, 217.23, Y_COIL)
    L(217.23, Y_0VI, 217.23, Y_0V)        # from internal 0V to external 0V
    L(217.28, 102.56, 216.94, 102.56)     # NK horizontal stub
    BLK('role aciklama', 217.23, 98.84)
    T("0.3", 217.08, 106.65, 0.14)
    T("0.0", 217.08, 105.25, 0.14)
    T("A2",  217.33, 99.55, 0.14)
    T("A1",  217.33, 100.05, 0.14)
    T("1X5", 216.66, 100.64, 0.12)   # cable ref (örnek position)
    T("7",   217.38, 100.64, 0.12)
    T("11",  217.38, 101.94, 0.12)
    T("10",  217.38, 103.04, 0.12)
    T("1X5", 216.66, 104.34, 0.12)   # cable ref (örnek position)
    T("6",   217.38, 104.34, 0.12)
    T(valve_label.upper(), 216.64, 101.74, 0.14)
    T("travel limit", 216.91, 101.78, 0.12)
    T("switch", 217.05, 101.93, 0.12)
    T("11K9", 216.99, 99.77, 0.14)
    T("OPEN", 217.02, 97.51, 0.14)
    T("O",    216.98, 102.65, 0.12)

    # ── K2 CLOSE coil rung (X=218.43) — 8 segments matching örnek ───────
    L(218.43, 104.46, 218.43, 105.13)
    L(218.43, 104.36, 218.43, 103.14)
    BLK('NK', 218.43, Y_NK)
    L(218.43, 103.04, 218.43, Y_NK)
    L(218.43, 102.24, 218.43, 102.04)
    L(218.43, 101.94, 218.43, 100.76)
    BLK('BOBIN', 218.43, Y_COIL)
    L(218.43, 100.66, 218.43, Y_COIL)
    L(218.43, Y_0VI, 218.43, Y_0V)
    L(218.48, 102.56, 218.14, 102.56)
    BLK('role aciklama', 218.43, 98.84)
    T("0.4", 218.88, 106.65, 0.14)
    T("0.1", 218.88, 105.25, 0.14)
    T("A2",  218.53, 99.55, 0.14)
    T("A1",  218.53, 100.05, 0.14)
    T("9",   218.58, 100.64, 0.12)
    T("14",  218.58, 101.94, 0.12)
    T("13",  218.58, 103.04, 0.12)
    T("8",   218.58, 104.34, 0.12)
    T("travel limit", 218.11, 101.81, 0.12)
    T("switch", 218.25, 101.96, 0.12)
    T("11K10", 218.13, 99.77, 0.14)
    T("CLOSE", 218.17, 97.51, 0.14)
    T("C",    218.17, 102.65, 0.12)

    # ── Additional DI/DQ channel labels (to right of K2) ─────────────────
    T("0.5", 220.08, 106.65, 0.14)
    T("0.2", 220.68, 105.25, 0.14)
    T("0.6", 221.28, 106.65, 0.14)
    T("0.3", 222.48, 105.25, 0.14)
    T("0.7", 222.48, 106.65, 0.14)
    T("1.0", 223.68, 106.65, 0.14)
    T("0.4", 224.28, 105.25, 0.14)
    T("OUTPUT", 224.4, 105.61, 0.16)
    T("INPUT",  224.44, 106.24, 0.16)
    T("1.1", 224.88, 106.65, 0.14)
    T("0.5", 226.08, 105.25, 0.14)
    T("1.2", 226.08, 106.65, 0.14)
    T("2L",  226.71, 105.25, 0.14)
    T("1.3", 227.28, 106.65, 0.14)
    T("0.6", 227.88, 105.25, 0.14)
    T("1.4", 228.48, 106.65, 0.14)
    T("0.7", 229.68, 105.25, 0.14)
    T("1.5", 229.68, 106.65, 0.14)
    T("1.0", 231.48, 105.25, 0.14)
    T("1.1", 233.28, 105.25, 0.14)

    # ═══════════════════════════════════════════════════════════════════════
    # TERMINAL / WIRING SECTION  (X ≈ 228–255)  — exact örnek coordinates
    # ═══════════════════════════════════════════════════════════════════════

    # ── OK power supply terminals + bus extension to them ─────────────────
    terminal(235.42, Y_PROT)           # L- terminal
    terminal(235.42, Y_LP)             # L+ terminal
    L(234.25, Y_PROT, 235.42, Y_PROT) # bus extension to OK terminal
    T("+",    235.3,  112.78, 0.16)
    T("-",    235.31, 112.42, 0.16)
    T("11/0", 235.67, 112.75, 0.14)
    T("11/0", 235.67, 112.36, 0.14)

    # PLC boundary near terminal section
    L(234.93, 105.13, 235.41, 105.13)
    L(234.93, 105.53, 235.56, 105.53)
    L(234.93, 106.53, 235.26, 106.53)
    L(234.93, 106.93, 235.41, 106.93)
    # Terminal section arcs (örnek ET4 entities)
    A(234.81, 105.58, 0.75)
    A(236.01, 106.48, 0.75)

    # External fire system label
    T("EXTERNAL FIRE FIGHTING SYSTEM", 243.44, 112.29, 0.18)
    T("CONTROL CABINET", 244.6, 111.89, 0.16)
    T("1X5", 245.53, 111.24, 0.14)

    # ── Wire (cable label) blocks + vertical connections ──────────────────
    cable_lbl(247.32, 107.8)
    T("7x1.5mm2",  247.22, 107.86, 0.16)
    T("1W5-3",     247.4,  106.89, 0.14)
    L(247.32, 107.8, 247.32, 110.21)   # up to upper bus
    L(247.32, 106.6, 247.32, 104.26)   # down to connector top bus

    cable_lbl(250.16, 107.8)
    T("4x2x0.75mm2", 250.06, 107.86, 0.16)
    T("1W5-4",       250.24, 106.89, 0.14)
    L(250.16, 107.8, 250.16, 110.21)
    L(250.16, 106.6, 250.16, 104.26)

    # ── Cable connector assembly body (X=246.57-251.67, Y=103.36-103.96) ─
    L(246.57, 103.36, 251.67, 103.36)   # bottom rail
    L(246.57, 103.96, 251.67, 103.96)   # top rail
    L(246.57, 103.36, 246.57, 103.96)   # left wall
    for _xd in (246.87, 247.17, 247.47, 247.77, 248.07, 248.37, 248.67,
                248.97, 249.27, 249.57, 249.87, 250.17, 250.47, 250.77,
                251.07, 251.37):
        L(_xd, 103.36, _xd, 103.96)
    L(251.67, 103.36, 251.67, 103.96)   # right wall

    # ── W1 cable group conductors (power: L1/L2/L3/PE/26/27) ─────────────
    # Top bus for W1 group (Y=104.26)
    L(247.92, 104.26, 246.72, 104.26)
    # Upper bus for W1 group (Y=110.21)
    L(247.92, 110.21, 246.72, 110.21)

    # Individual W1 conductors (top stubs + upper stubs)
    for _xc, _diag_coords in [
        (246.72, (246.72, 102.59, 246.84, 102.44)),
        (247.02, None),
        (247.32, (247.32, 102.59, 247.21, 102.44)),
        (247.62, None),
        (247.92, None),
    ]:
        L(_xc, 104.11, _xc, 103.96)
        L(_xc, 104.26, _xc, 104.11)
        L(_xc, 110.36, _xc, 110.21)
        L(_xc, 110.51, _xc, 110.36)
        if _diag_coords:
            L(*_diag_coords)
    # W1 straight bottom conductors
    L(247.02, 103.36, 247.02, 102.49)
    L(247.62, 101.81, 247.62, 103.36)
    L(247.62, 101.81, 247.92, 101.81)
    L(247.92, 101.96, 247.92, 101.81)
    L(247.92, 103.36, 247.92, 102.56)   # to NK
    L(246.72, 103.36, 246.72, 102.59)
    L(247.32, 103.36, 247.32, 102.59)
    # Extra dividers W1 area
    L(248.07, 103.36, 247.77, 103.36)

    # ── NK/NA travel limit contacts (field side) ───────────────────────────
    BLK("NK", 249.12, 102.56)           # OPEN NC limit
    BLK("NK", 250.32, 102.56)           # CLOSE NC limit
    BLK("NA", 250.92, 102.56)           # OPEN NO limit
    BLK("NA", 251.52, 102.56)           # CLOSE NO limit
    L(248.92, 102.26, 249.17, 102.26)
    L(250.12, 102.26, 250.37, 102.26)
    L(250.72, 102.26, 250.86, 102.26)
    L(251.32, 102.26, 251.46, 102.26)

    # ── W2 cable group conductors (signal: limit switches, thermostat) ────
    # Top bus for W2 group (Y=104.26)
    L(251.52, 104.26, 248.81, 104.26)
    # Upper bus for W2 group (Y=110.21)
    L(251.21, 110.21, 249.11, 110.21)

    # W2 conductors with top stubs, upper stubs, and bottom connections
    for _xc, _bot_y, _has_upper in [
        (248.82, 101.81, True),
        (249.12, 102.56, False),
        (250.02, 101.81, True),
        (250.32, 102.56, False),
        (250.62, 101.81, False),
        (250.92, 102.56, False),
        (251.22, 101.81, False),
        (251.52, 102.56, True),
    ]:
        L(_xc, 104.11, _xc, 103.96)
        L(_xc, 104.26, _xc, 104.11)
        if _bot_y == 101.81:
            L(_xc, 101.81, _xc, 103.36)
        else:
            L(_xc, 101.96, _xc, 101.81)
            L(_xc, 103.36, _xc, 102.56)
    for _bot_x, _nxt_x in [(248.82, 249.12), (250.02, 250.32), (250.62, 250.92), (251.22, 251.52)]:
        L(_bot_x, 101.81, _nxt_x, 101.81)

    # Upper stubs for W2 group
    for _xc in (249.11, 249.41, 249.71, 250.01, 250.31, 250.61, 250.91, 251.21):
        L(_xc, 110.36, _xc, 110.21)
        L(_xc, 110.51, _xc, 110.36)

    # W2 dividers
    for _xd in (249.27, 249.57, 249.87, 250.17, 250.47, 250.77, 251.07, 251.37, 251.67):
        pass   # already drawn as connector body dividers above

    # ── Terminal labels — W1 power cable ──────────────────────────────────
    for _x, _tn, _tb, _tt, _tu in [
        (246.72, "1",  "1", "1",  "14"),
        (247.02, "2",  "2", "2",  "15"),
        (247.32, "3",  "3", "3",  "16"),
        (247.62, "4",  "26","4",  "17"),
        (247.92, "5",  "27","5",  "18"),
    ]:
        T(_tn,  _x - 0.04, 104.31, 0.12)
        T(_tb,  _x - 0.04, 103.59, 0.12)
        T(_tt,  _x - 0.04, 110.06, 0.12)
        T(_tu,  _x - 0.04, 110.75, 0.12)

    # W1 cable labels
    T("400VAC", 246.77, 101.42, 0.14)
    T("50Hz",   246.85, 101.28, 0.14)
    T("3P~",    246.89, 101.15, 0.14)
    T("3P~",    246.89, 101.94, 0.14)
    T("M",      246.91, 102.14, 0.14)
    T("400VAC", 247.68, 101.07, 0.14)
    T("heater", 247.82, 101.11, 0.14)
    T(valve_label.upper() + " W1", 247.8, 100.49, 0.14)

    # ── Terminal labels — W2 signal cable ─────────────────────────────────
    # Bottom (104.31, 103.59) use conductor X - 0.04
    # Upper (110.06, 110.75) use upper-stub X positions (249.11 series) - different offset
    for (_bx, _tx), (_tn, _tb, _tt, _tu) in zip(
        zip([248.82, 249.12, 250.02, 250.32, 250.62, 250.92, 251.22, 251.52],
            [249.11, 249.41, 249.71, 250.01, 250.31, 250.61, 250.91, 251.21]),
        [("1","10","1","19"), ("2","11","2","20"), ("3","13","4","21"),
         ("4","14","5","22"), ("5","20","6","23"), ("6","22","7","24"),
         ("7","23","8","25"), ("8","25",None,"26")]
    ):
        T(_tn,  _bx - 0.04, 104.31, 0.12)
        T(_tb,  _bx - 0.04, 103.59, 0.12)
        if _tt:
            T(_tt, _tx - 0.03, 110.06, 0.12)
        T(_tu,  _tx - 0.10, 110.75, 0.12)

    # Travel limit and contact labels (field side)
    T("travel limit", 248.88, 100.95, 0.12)
    T("switch",       249.02, 101.12, 0.12)
    T("opened",       249.14, 101.09, 0.12)
    T("nc",           248.92, 101.96, 0.12)
    T("1",            249.17, 101.96, 0.12)
    T("2",            249.17, 102.51, 0.12)
    T("travel limit", 250.08, 100.95, 0.12)
    T("switch",       250.22, 101.12, 0.12)
    T("closed",       250.34, 101.11, 0.12)
    T("nc",           250.12, 101.96, 0.12)
    T("1",            250.37, 101.96, 0.12)
    T("2",            250.37, 102.51, 0.12)
    T("extra limit",  250.68, 100.96, 0.12)
    T("switch",       250.82, 101.12, 0.12)
    T("opened",       250.94, 101.09, 0.12)
    T("nc",           250.72, 101.96, 0.12)
    T("1",            250.97, 101.96, 0.12)
    T("3",            250.97, 102.51, 0.12)
    T("extra limit",  251.28, 100.96, 0.12)
    T("switch",       251.42, 101.12, 0.12)
    T("closed",       251.54, 101.11, 0.12)
    T("nc",           251.32, 101.96, 0.12)
    T("1",            251.57, 101.96, 0.12)
    T("3",            251.57, 102.51, 0.12)

    # ═══════════════════════════════════════════════════════════════════════
    # HEADER
    # ═══════════════════════════════════════════════════════════════════════
    T(f"{valve_label.upper()}  CONTROL CABINET", 208.0, Y_TOP + 0.35, 0.35)
    T(project_name + ("  |  " + ship_name if ship_name else ""), 195.0, Y_TOP - 0.30, 0.22)
    T(f"DWG: {drawing_no}  |  Date: {date_str}", 230.0, Y_TOP - 0.30, 0.20)
    if company:
        T(company, 195.0, Y_TOP + 0.35, 0.22)

    # ═══════════════════════════════════════════════════════════════════════
    # LAYOUT VIEWPORT  –  use "ISO A1 Title Block" from template (already set)
    # ═══════════════════════════════════════════════════════════════════════
    try:
        doc.Regen(1)
    except Exception:
        pass

    try:
        lo = doc.Layouts.Item("ISO A1 Title Block")
        doc.ActiveLayout = lo
        time.sleep(0.4)
        doc.SendCommand("MSPACE\n")
        time.sleep(0.3)
        doc.SendCommand(f"ZOOM\nW\n193.0,{Y_BOT - 1.0}\n262.0,{Y_TOP + 1.5}\n")
        time.sleep(0.3)
        doc.SendCommand("PSPACE\n")
        time.sleep(0.2)
    except Exception:
        pass

    # Save: use SaveAs to output_path (works even if doc was created as untitled)
    try:
        doc.SaveAs(output_path)
    except Exception:
        try:
            doc.Save()
        except Exception:
            pass

    return (
        f"OK: '{valve_label}' [{aq_key}] saved to {output_path}\n"
        f"Motor prot: {rv_part}  Ir={rv_range}A  set={rv_set}A  + 3RV2901-1D aux\n"
        f"Power MCB: {main_mcb} ({P}F1)  Ctrl MCB: {ctrl_mcb} ({P}F2)\n"
        f"Contactors: 3TG1010-0BB4 24VDC ({P}K1 OPEN, {P}K2 CLOSE)\n"
        f"Phase monitor: 3UG4512-1AR20 ({P}A1)\n"
        f"PLC: {plc_model}  DI I0.0-I0.6  DQ Q0.0-Q0.1\n"
        f"AQ specs: {torque}Nm {t_s}s/90 {kW}kW In={In}A Istart={Istart}A\n"
        f"Terminals: {P}X1 -> AQ (1-3/PE/26/27/10-15/40-41)\n"
        f"Cables: W1=7x1.5mm2  W2=4x2x0.75mm2 shielded\n"
        f"Supported AQ: AQ5/10/15/25/30/50/80/150/280/430/610/830/1000"
    )


def main():
    log.info("Starting AutoCAD MCP server (stdio transport)")
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()