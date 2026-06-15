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
# ProgID "ZWCAD.Application", so we try it too — the drawing tools below work the
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
        # raises "RPC server unavailable" — so drop it and reconnect below.
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
def add_text(text: str, x: float, y: float, height: float = 2.5) -> str:
    """Place single-line text at (x, y) with the given text height."""
    ms = _model_space()
    ms.AddText(text, _point(x, y), float(height))
    return f"Text '{text}' placed at ({x}, {y})."


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
def insert_block(block_name: str, x: float, y: float,
                 scale: float = 1.0, rotation_deg: float = 0.0) -> str:
    """Insert a block reference by name at (x, y) with uniform scale and rotation (degrees)."""
    import math
    ms = _model_space()
    ms.InsertBlock(
        _point(x, y), block_name,
        float(scale), float(scale), float(scale),
        math.radians(float(rotation_deg)),
    )
    return f"Block '{block_name}' inserted at ({x}, {y})."


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
            f"at {angle_deg}°, clipped to the outline."
        )
    except Exception as e:
        return (
            f"Could not create the hatch: {e}. The boundary must be a single closed shape, or "
            f"connected segments that form one closed loop. Use list_entities to find the right "
            f"index, and make sure the outline is actually closed."
        )


def main():
    log.info("Starting AutoCAD MCP server (stdio transport)")
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()