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
# NOTE: exact terminal NUMBERS come from the order-specific Bernard wiring sheet
# inside the unit's cover; placeholders are marked [t?] and must be confirmed.
# ===========================================================================

@mcp.tool()
def draw_aq_valve_control(origin_x: float = 0.0, origin_y: float = 0.0,
                          size: str = "AQ20", tag_prefix: str = "") -> str:
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
    NOTE: terminal numbers follow the order-specific Bernard wiring sheet --
    placeholders [t?] here must be confirmed against that sheet."""
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
    text("[t? / t?]", 265, 225, 2.4, "center")

    mtext("POWER NOTES (Bernard AQ SWITCH, SUG_17003 Ch.11.1):\\n"
          "1. Reversing pair -KM1 (OPEN/CCW=S1) / -KM2 (CLOSE/CW=S2). -KM2 swaps "
          "L1<->L3. SWITCH type has NO auto phase correction -> reverse by swapping "
          "two phases; phase sequence matters.\\n"
          "2. -K0 (3UG) phase-sequence/loss permissive across the incoming bus.\\n"
          "3. -Q2 (3RV) = short-circuit/backup. Motor thermal protector (thermostat, NC) "
          "is primary thermal protection and is wired into the coil control (see control "
          "sheet), per SUG_17003 7.3.\\n"
          "4. Heater (D) 230VAC kept permanently energised (anti-condensation).\\n"
          "5. Terminal numbers [t?] follow this unit's Bernard order wiring sheet - confirm.",
          40, 290, 175, 3.0)

    # =====================================================================
    # ACTUATOR TERMINAL STRIP (centre)
    # =====================================================================
    text(f"{size} ACTUATOR TERMINALS (per Bernard order wiring sheet)", 300, 200, 4)
    rect(300, 60, 470, 185)
    strip = [
        ("U / V / W", "3~ motor (E)  ->  -Q2 load"),
        ("Th1 / Th2", "Motor thermostat (NC)  -> coil string + PLC"),
        ("S2 close LS (B)", "CW travel limit -> stops -KM2 coil"),
        ("S1 open LS (B)", "CCW travel limit -> stops -KM1 coil"),
        ("CL torque (A)", "Close torque -> stops -KM2 coil + PLC"),
        ("OP torque (A)", "Open torque -> stops -KM1 coil + PLC"),
        ("CL sig (C)", "Closed signalling -> PLC DI"),
        ("OP sig (C)", "Open signalling -> PLC DI"),
        ("Heater (D)", "230 VAC anti-condensation"),
        ("PE", "Internal ground post"),
    ]
    yy = 178
    for term, desc in strip:
        text(term, 305, yy, 2.7)
        text(desc, 365, yy, 2.4)
        yy -= 12

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
    text(f"{pre}-F  THERMOSTAT (NC)", 525, 547, 2.4)
    line(556, 540, 575, 540)
    dot(575, 540)
    line(575, 540, 575, 500)

    # OPEN rung -> -KM1
    yO = 500
    line(575, yO, 590, yO)
    contact(598, yO, "nc"); text("S1 OPEN LIMIT", 588, yO + 5, 2.1)
    line(606, yO, 618, yO)
    contact(626, yO, "nc"); text("OPEN TORQUE", 616, yO + 5, 2.1)
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
    contact(598, yC, "nc"); text("S2 CLOSE LIMIT", 588, yC + 5, 2.1)
    line(606, yC, 618, yC)
    contact(626, yC, "nc"); text("CLOSE TORQUE", 616, yC + 5, 2.1)
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
          "5. Terminal numbers per this unit's Bernard order wiring sheet - confirm.",
          500, 130, 290, 3.0)

    return (f"Drew the datasheet-wired Bernard {size} valve control schematic "
            f"(power + actuator terminals + PLC/HMI, hardware-integrated limit/"
            f"torque/thermostat) at origin ({ox}, {oy}) on layer EFF-VALVE.")


def main():
    log.info("Starting AutoCAD MCP server (stdio transport)")
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()