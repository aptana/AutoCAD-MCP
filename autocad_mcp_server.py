from __future__ import annotations

import json
import os
from typing import Any, Iterable

import pythoncom
import win32com.client
from win32com.client import VARIANT
from mcp.server.fastmcp import FastMCP


mcp = FastMCP("autocad")


def get_autocad() -> Any:
    """Return the running AutoCAD 2021 COM application."""
    pythoncom.CoInitialize()
    try:
        return win32com.client.GetActiveObject("AutoCAD.Application.24")
    except Exception as error:
        raise RuntimeError(
            "AutoCAD Civil 3D 2021 tidak ditemukan. Pastikan AutoCAD sedang terbuka."
        ) from error


def active_document() -> Any:
    autocad = get_autocad()
    if autocad.Documents.Count == 0:
        raise RuntimeError("Belum ada drawing AutoCAD yang terbuka.")
    return autocad.ActiveDocument


def point(x: float, y: float, z: float = 0.0) -> Any:
    """Create the SAFEARRAY point type required by AutoCAD COM methods."""
    return VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        (float(x), float(y), float(z)),
    )


def radians(degrees: float) -> float:
    return float(degrees) * 3.141592653589793 / 180.0


@mcp.tool()
def get_active_document() -> str:
    """Return the name of the active AutoCAD document."""
    document = active_document()
    return f"{document.Name} | {document.FullName}"


@mcp.tool()
def list_documents() -> list[str]:
    """List all open AutoCAD drawings."""
    autocad = get_autocad()
    return [str(autocad.Documents.Item(index).Name) for index in range(autocad.Documents.Count)]


@mcp.tool()
def activate_document(name: str) -> str:
    """Activate an already open drawing by file name."""
    autocad = get_autocad()
    for index in range(autocad.Documents.Count):
        document = autocad.Documents.Item(index)
        if document.Name.lower() == name.lower():
            document.Activate()
            return f"Drawing aktif: {document.Name}"
    raise FileNotFoundError(f"Drawing terbuka tidak ditemukan: {name}")


@mcp.tool()
def open_document(path: str) -> str:
    """Open a DWG/DWT file and make it active."""
    full_path = os.path.abspath(path)
    if not os.path.isfile(full_path):
        raise FileNotFoundError(full_path)
    document = get_autocad().Documents.Open(full_path)
    document.Activate()
    return f"Dibuka: {document.Name}"


@mcp.tool()
def new_document(template_path: str = "") -> str:
    """Create a new drawing, optionally from a DWT template."""
    autocad = get_autocad()
    document = autocad.Documents.Add(template_path) if template_path else autocad.Documents.Add()
    document.Activate()
    return f"Drawing baru: {document.Name}"


@mcp.tool()
def save_document(path: str = "") -> str:
    """Save the active drawing, optionally to a new DWG path."""
    document = active_document()
    if path:
        document.SaveAs(os.path.abspath(path))
    else:
        document.Save()
    return f"Tersimpan: {document.FullName}"


@mcp.tool()
def close_document(save_changes: bool = True) -> str:
    """Close the active drawing, optionally saving changes first."""
    document = active_document()
    name = document.Name
    document.Close(bool(save_changes))
    return f"Drawing ditutup: {name}"


@mcp.tool()
def set_view(mode: str = "extents") -> str:
    """Set the active view: extents, plan, or world_ucs."""
    document = active_document()
    normalized = mode.lower().strip()
    if normalized == "extents":
        document.SendCommand("_.ZOOM\n_E\n")
    elif normalized in {"plan", "world_ucs"}:
        document.SendCommand("_.UCS\n_W\n_.PLAN\n_W\n")
    else:
        raise ValueError("mode harus extents, plan, atau world_ucs.")
    document.Regen(0)
    return f"View diubah: {normalized}"


@mcp.tool()
def create_layer(name: str, color_index: int = 7) -> str:
    """Create or activate a layer. AutoCAD color index is 1-255."""
    document = active_document()
    try:
        layer = document.Layers.Item(name)
    except Exception:
        layer = document.Layers.Add(name)
    layer.Color = max(1, min(255, int(color_index)))
    document.ActiveLayer = layer
    return f"Layer aktif: {layer.Name}"


@mcp.tool()
def create_line(
    start_x: float,
    start_y: float,
    end_x: float,
    end_y: float,
) -> str:
    """Create a line in the active AutoCAD drawing using WCS coordinates."""
    document = active_document()
    entity = document.ModelSpace.AddLine(point(start_x, start_y), point(end_x, end_y))
    document.Regen(0)
    return f"Line dibuat, panjang {entity.Length:.3f}."


@mcp.tool()
def create_circle(center_x: float, center_y: float, radius: float, layer: str = "") -> str:
    """Create a circle in WCS coordinates."""
    if radius <= 0:
        raise ValueError("radius harus lebih besar dari 0.")
    document = active_document()
    entity = document.ModelSpace.AddCircle(point(center_x, center_y), float(radius))
    if layer:
        entity.Layer = layer
    document.Regen(0)
    return f"Circle dibuat, radius {radius:.3f}."


@mcp.tool()
def create_rectangle(min_x: float, min_y: float, max_x: float, max_y: float, layer: str = "") -> str:
    """Create a closed rectangular lightweight polyline in WCS coordinates."""
    if max_x <= min_x or max_y <= min_y:
        raise ValueError("max_x/max_y harus lebih besar dari min_x/min_y.")
    document = active_document()
    coordinates = (min_x, min_y, max_x, min_y, max_x, max_y, min_x, max_y)
    entity = document.ModelSpace.AddLightWeightPolyline(coordinates)
    entity.Closed = True
    if layer:
        entity.Layer = layer
    document.Regen(0)
    return f"Rectangle dibuat: {max_x - min_x:g} x {max_y - min_y:g}."


@mcp.tool()
def create_polyline(points: list[float], closed: bool = False, layer: str = "") -> str:
    """Create a 2D polyline from [x1,y1,x2,y2,...] WCS coordinates."""
    if len(points) < 4 or len(points) % 2 != 0:
        raise ValueError("points harus berisi minimal dua pasangan x,y.")
    document = active_document()
    entity = document.ModelSpace.AddLightWeightPolyline(tuple(float(value) for value in points))
    entity.Closed = bool(closed)
    if layer:
        entity.Layer = layer
    document.Regen(0)
    return f"Polyline dibuat dengan {len(points) // 2} titik."


@mcp.tool()
def add_text(text: str, x: float, y: float, height: float = 2.5, rotation_degrees: float = 0.0, layer: str = "") -> str:
    """Add single-line text in WCS coordinates."""
    if height <= 0:
        raise ValueError("height harus lebih besar dari 0.")
    document = active_document()
    entity = document.ModelSpace.AddText(text, point(x, y), float(height))
    entity.Rotation = radians(rotation_degrees)
    if layer:
        entity.Layer = layer
    document.Regen(0)
    return f"Text dibuat: {text}"


@mcp.tool()
def list_blocks() -> list[str]:
    """List block definitions in the active drawing."""
    document = active_document()
    return [str(document.Blocks.Item(index).Name) for index in range(document.Blocks.Count)]


@mcp.tool()
def list_modelspace_entities(limit: int = 200) -> list[str]:
    """List basic information and handles for model-space entities."""
    if limit < 1:
        raise ValueError("limit harus lebih besar dari 0.")
    document = active_document()
    result = []
    for index, entity in enumerate(document.ModelSpace):
        if index >= limit:
            break
        result.append(f"{entity.ObjectName} | handle={entity.Handle} | layer={entity.Layer}")
    return result


@mcp.tool()
def insert_block(name: str, x: float, y: float, z: float = 0.0, scale: float = 1.0, rotation_degrees: float = 0.0) -> str:
    """Insert an existing block definition at a WCS point."""
    if scale <= 0:
        raise ValueError("scale harus lebih besar dari 0.")
    document = active_document()
    try:
        document.Blocks.Item(name)
    except Exception as error:
        raise ValueError(f"Block tidak ditemukan: {name}") from error
    document.ModelSpace.InsertBlock(
        point(x, y, z), name, float(scale), float(scale), float(scale), radians(rotation_degrees)
    )
    document.Regen(0)
    return f"Block {name} ditempatkan pada ({x}, {y}, {z})."


def bounding_box(entity: Any) -> tuple[float, float, float, float, float, float]:
    lower, upper = entity.GetBoundingBox()
    return (
        float(lower[0]), float(lower[1]), float(lower[2]),
        float(upper[0]), float(upper[1]), float(upper[2]),
    )


def find_block_reference(name: str) -> Any:
    document = active_document()
    for entity in document.ModelSpace:
        if entity.ObjectName == "AcDbBlockReference" and entity.Name.lower() == name.lower():
            return entity
    raise ValueError(f"Block reference tidak ditemukan di Model Space: {name}")


@mcp.tool()
def get_entity_bounding_box(handle: str) -> dict[str, float]:
    """Return the WCS bounding box of an entity or block reference."""
    lower_x, lower_y, lower_z, upper_x, upper_y, upper_z = bounding_box(entity_by_handle(handle))
    return {
        "min_x": lower_x,
        "min_y": lower_y,
        "min_z": lower_z,
        "max_x": upper_x,
        "max_y": upper_y,
        "max_z": upper_z,
        "width": upper_x - lower_x,
        "height": upper_y - lower_y,
    }


@mcp.tool()
def arrange_blocks_with_left_connector(
    right_block_name: str,
    left_block_name: str,
    x: float = 0.0,
    y: float = 0.0,
    right_block_rotation_degrees: float = 90.0,
    connector_length: float = 25.0,
    scale: float = 1.0,
) -> dict[str, Any]:
    """Place a rotated right block, a left-pointing connector, and a left block edge-to-edge."""
    if connector_length <= 0 or scale <= 0:
        raise ValueError("connector_length dan scale harus lebih besar dari 0.")
    document = active_document()
    try:
        document.Blocks.Item(right_block_name)
        document.Blocks.Item(left_block_name)
    except Exception as error:
        raise ValueError("Salah satu block definition tidak ditemukan.") from error

    right_block = document.ModelSpace.InsertBlock(
        point(x, y), right_block_name, float(scale), float(scale), float(scale), radians(right_block_rotation_degrees)
    )
    document.Regen(0)
    right_min_x, right_min_y, _, right_max_x, right_max_y, _ = bounding_box(right_block)
    center_y = (right_min_y + right_max_y) / 2.0
    connector_start_x = right_min_x
    connector_end_x = connector_start_x - float(connector_length)
    connector = document.ModelSpace.AddLine(
        point(connector_end_x, center_y), point(connector_start_x, center_y)
    )

    left_block = document.ModelSpace.InsertBlock(
        point(connector_end_x, center_y), left_block_name, float(scale), float(scale), float(scale), 0.0
    )
    document.Regen(0)
    left_min_x, left_min_y, _, left_max_x, left_max_y, _ = bounding_box(left_block)
    left_center_y = (left_min_y + left_max_y) / 2.0
    left_block.Move(
        point(0.0, 0.0),
        point(connector_end_x - left_max_x, center_y - left_center_y),
    )
    document.Regen(0)
    document.SendCommand("_.ZOOM\n_E\n")
    return {
        "right_block_handle": str(right_block.Handle),
        "connector_handle": str(connector.Handle),
        "left_block_handle": str(left_block.Handle),
        "right_block_bbox": dict(zip(("min_x", "min_y", "min_z", "max_x", "max_y", "max_z"), bounding_box(right_block))),
        "connector": {"start_x": connector_end_x, "end_x": connector_start_x, "y": center_y, "length": connector_length},
        "left_block_bbox": dict(zip(("min_x", "min_y", "min_z", "max_x", "max_y", "max_z"), bounding_box(left_block))),
    }


@mcp.tool()
def insert_external_block(path: str, x: float, y: float, z: float = 0.0, scale: float = 1.0, rotation_degrees: float = 0.0) -> str:
    """Insert a DWG file as a block reference in the active drawing."""
    full_path = os.path.abspath(path)
    if not os.path.isfile(full_path):
        raise FileNotFoundError(full_path)
    if scale <= 0:
        raise ValueError("scale harus lebih besar dari 0.")
    document = active_document()
    entity = document.ModelSpace.InsertBlock(
        point(x, y, z), full_path, float(scale), float(scale), float(scale), radians(rotation_degrees)
    )
    document.Regen(0)
    return f"Block eksternal ditempatkan, handle={entity.Handle}."


@mcp.tool()
def erase_entity(handle: str) -> str:
    """Erase one entity by its AutoCAD handle after it has been identified."""
    document = active_document()
    entity = document.HandleToObject(handle)
    entity.Delete()
    document.Regen(0)
    return f"Entity dihapus: {handle}"


def entity_by_handle(handle: str) -> Any:
    return active_document().HandleToObject(handle)


@mcp.tool()
def get_entity_properties(handle: str) -> dict[str, Any]:
    """Read common properties from an entity handle."""
    entity = entity_by_handle(handle)
    result: dict[str, Any] = {
        "handle": str(entity.Handle),
        "object_name": str(entity.ObjectName),
        "layer": str(entity.Layer),
        "color": int(entity.Color),
        "linetype": str(entity.Linetype),
        "lineweight": int(entity.Lineweight),
    }
    for name in ("Length", "Radius", "Area", "Rotation", "TextString"):
        try:
            value = getattr(entity, name)
            result[name.lower()] = str(value)
        except Exception:
            pass
    return result


@mcp.tool()
def set_entity_properties(handle: str, layer: str = "", color_index: int = 0, linetype: str = "") -> str:
    """Change common properties of an entity; 0 leaves color unchanged."""
    entity = entity_by_handle(handle)
    if layer:
        entity.Layer = layer
    if color_index:
        entity.Color = max(1, min(255, int(color_index)))
    if linetype:
        entity.Linetype = linetype
    active_document().Regen(0)
    return f"Properti entity diubah: {handle}"


@mcp.tool()
def transform_entities(handles: list[str], operation: str, base_x: float = 0.0, base_y: float = 0.0, value_x: float = 0.0, value_y: float = 0.0, angle_degrees: float = 0.0, scale: float = 1.0) -> list[str]:
    """Move, copy, rotate, mirror, scale, or offset entities by handle."""
    if not handles:
        raise ValueError("handles tidak boleh kosong.")
    document = active_document()
    base = point(base_x, base_y)
    target = point(value_x, value_y)
    normalized = operation.lower().strip()
    results: list[str] = []
    for handle in handles:
        entity = document.HandleToObject(handle)
        if normalized == "move":
            entity.Move(base, target)
            results.append(handle)
        elif normalized == "copy":
            results.append(str(entity.Copy().Handle))
            copied = document.HandleToObject(results[-1])
            copied.Move(base, target)
        elif normalized == "rotate":
            entity.Rotate(base, radians(angle_degrees))
            results.append(handle)
        elif normalized == "mirror":
            mirrored = entity.Mirror(base, target)
            results.append(str(mirrored.Handle))
        elif normalized == "scale":
            if scale <= 0:
                raise ValueError("scale harus lebih besar dari 0.")
            entity.ScaleEntity(base, float(scale))
            results.append(handle)
        elif normalized == "offset":
            offset_entities = entity.Offset(float(value_x))
            results.extend(str(offset_entities.Item(index).Handle) for index in range(offset_entities.Count))
        else:
            raise ValueError("operation harus move, copy, rotate, mirror, scale, atau offset.")
    document.Regen(0)
    return results


@mcp.tool()
def create_aligned_dimension(x1: float, y1: float, x2: float, y2: float, text_x: float, text_y: float, layer: str = "") -> str:
    """Create an aligned dimension in model space."""
    document = active_document()
    entity = document.ModelSpace.AddDimAligned(point(x1, y1), point(x2, y2), point(text_x, text_y))
    if layer:
        entity.Layer = layer
    document.Regen(0)
    return f"Dimension dibuat: handle={entity.Handle}"


@mcp.tool()
def create_linear_dimension(x1: float, y1: float, x2: float, y2: float, text_x: float, text_y: float, rotation_degrees: float = 0.0, layer: str = "") -> str:
    """Create a rotated/linear dimension in model space."""
    document = active_document()
    entity = document.ModelSpace.AddDimRotated(
        point(x1, y1), point(x2, y2), point(text_x, text_y), radians(rotation_degrees)
    )
    if layer:
        entity.Layer = layer
    document.Regen(0)
    return f"Dimension linear dibuat: handle={entity.Handle}"


@mcp.tool()
def set_active_layout(name: str) -> str:
    """Switch the active paper-space layout."""
    document = active_document()
    try:
        layout = document.Layouts.Item(name)
    except Exception as error:
        raise ValueError(f"Layout tidak ditemukan: {name}") from error
    document.ActiveLayout = layout
    return f"Layout aktif: {layout.Name}"


@mcp.tool()
def list_layouts() -> list[str]:
    """List layouts in the active drawing."""
    document = active_document()
    return [str(document.Layouts.Item(index).Name) for index in range(document.Layouts.Count)]


@mcp.tool()
def plot_active_layout_to_pdf(path: str) -> str:
    """Plot the active layout to a PDF file using its current page setup."""
    full_path = os.path.abspath(path)
    os.makedirs(os.path.dirname(full_path), exist_ok=True)
    document = active_document()
    document.Plot.PlotToFile(full_path, document.ActiveLayout.ConfigName)
    return f"PDF dibuat: {full_path}"


@mcp.tool()
def attach_xref(path: str, name: str, x: float = 0.0, y: float = 0.0, z: float = 0.0, scale: float = 1.0, rotation_degrees: float = 0.0) -> str:
    """Attach an external DWG as an Xref."""
    full_path = os.path.abspath(path)
    if not os.path.isfile(full_path):
        raise FileNotFoundError(full_path)
    if scale <= 0:
        raise ValueError("scale harus lebih besar dari 0.")
    document = active_document()
    entity = document.ModelSpace.AttachExternalReference(
        full_path, name, point(x, y, z), float(scale), float(scale), float(scale), radians(rotation_degrees), True
    )
    document.Regen(0)
    return f"Xref terpasang: {name}, handle={entity.Handle}"


@mcp.tool()
def import_geojson(
    path: str,
    layer_field: str = "layer",
    block_field: str = "block",
    label_field: str = "",
    coordinate_scale: float = 1.0,
    offset_x: float = 0.0,
    offset_y: float = 0.0,
) -> dict[str, Any]:
    """Import GeoJSON features from QGIS as AutoCAD lines, polylines, blocks, and labels."""
    full_path = os.path.abspath(path)
    if not os.path.isfile(full_path):
        raise FileNotFoundError(full_path)
    if coordinate_scale == 0:
        raise ValueError("coordinate_scale tidak boleh 0.")
    with open(full_path, "r", encoding="utf-8") as source:
        payload = json.load(source)
    if payload.get("type") != "FeatureCollection":
        raise ValueError("File harus berupa GeoJSON FeatureCollection.")

    document = active_document()
    created: dict[str, int] = {"points": 0, "lines": 0, "polygons": 0, "blocks": 0, "labels": 0}
    handles: list[str] = []

    def convert_coordinate(values: list[float]) -> tuple[float, float, float]:
        z = float(values[2]) if len(values) > 2 else 0.0
        return point(
            float(values[0]) * coordinate_scale + offset_x,
            float(values[1]) * coordinate_scale + offset_y,
            z * coordinate_scale,
        )

    def apply_layer(entity: Any, properties: dict[str, Any]) -> None:
        layer_name = properties.get(layer_field, "") if layer_field else ""
        if layer_name:
            try:
                document.Layers.Item(str(layer_name))
            except Exception:
                document.Layers.Add(str(layer_name))
            entity.Layer = str(layer_name)

    def add_label(properties: dict[str, Any], coordinates: tuple[float, float, float]) -> None:
        if not label_field or properties.get(label_field) in (None, ""):
            return
        label = document.ModelSpace.AddText(str(properties[label_field]), coordinates, 2.5)
        apply_layer(label, properties)
        handles.append(str(label.Handle))
        created["labels"] += 1

    for feature in payload["features"]:
        geometry = feature.get("geometry") or {}
        properties = feature.get("properties") or {}
        geometry_type = geometry.get("type")
        coordinates = geometry.get("coordinates")
        if not coordinates:
            continue
        block_name = properties.get(block_field, "") if block_field else ""
        if geometry_type == "Point":
            insertion = convert_coordinate(coordinates)
            if block_name:
                try:
                    document.Blocks.Item(str(block_name))
                except Exception as error:
                    raise ValueError(f"Block tidak ditemukan: {block_name}") from error
                entity = document.ModelSpace.InsertBlock(insertion, str(block_name), 1.0, 1.0, 1.0, 0.0)
                apply_layer(entity, properties)
                created["blocks"] += 1
            else:
                entity = document.ModelSpace.AddPoint(insertion)
                apply_layer(entity, properties)
                created["points"] += 1
            handles.append(str(entity.Handle))
            add_label(properties, insertion)
        elif geometry_type in {"LineString", "Polygon"}:
            rings = [coordinates] if geometry_type == "LineString" else coordinates
            for ring in rings:
                flat: list[float] = []
                for coordinate in ring:
                    converted = convert_coordinate(coordinate)
                    flat.extend((converted[0], converted[1]))
                entity = document.ModelSpace.AddLightWeightPolyline(tuple(flat))
                entity.Closed = geometry_type == "Polygon"
                apply_layer(entity, properties)
                handles.append(str(entity.Handle))
                created["lines" if geometry_type == "LineString" else "polygons"] += 1

    document.Regen(0)
    return {"path": full_path, "created": created, "handles": handles}


@mcp.tool()
def send_autocad_command(command: str) -> str:
    """Send one explicitly requested native AutoCAD command sequence."""
    if not command.strip():
        raise ValueError("command tidak boleh kosong.")
    document = active_document()
    document.SendCommand(command.rstrip() + "\n")
    return "Perintah AutoCAD dikirim."


if __name__ == "__main__":
    mcp.run()
