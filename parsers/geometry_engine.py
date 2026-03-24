"""
geometry_engine.py
Builds room polygons from classified wall entities using Shapely polygonize.
Also detects doors and windows by proximity to wall openings.
"""
from __future__ import annotations
import math
from typing import Any


def detect_rooms(dxf_path: str, min_area_m2: float = 1.0) -> list[dict]:
    """
    Full pipeline: parse → classify → polygonize walls → label rooms.
    Returns list of room dicts:
      {"id", "label", "area_m2", "centroid_x", "centroid_y",
       "points", "doors_nearby", "windows_nearby"}
    """
    from parsers.dxf_parser import parse_dxf
    from parsers.element_classifier import classify_all_layers

    data = parse_dxf(dxf_path)
    layer_types = classify_all_layers(data["layers"])
    uf = data["unit_factor"]

    rooms = (
        _rooms_from_mahal_blocks(dxf_path, uf, min_area_m2)
        or _rooms_from_hatch(dxf_path, uf, min_area_m2)
        or _rooms_from_polygonize(data["entities"], layer_types, uf, min_area_m2)
        or _rooms_from_closed_polylines(data["entities"], layer_types, uf, min_area_m2)
    )

    doors   = _collect_elements(data["entities"], layer_types, "doors")
    windows = _collect_elements(data["entities"], layer_types, "windows")

    for room in rooms:
        cx, cy = room["centroid_x"], room["centroid_y"]
        thresh = math.sqrt(room["area_m2"] / uf) * 1.2 if uf > 0 else 5000
        room["doors_nearby"]   = _nearby(doors, cx, cy, thresh)
        room["windows_nearby"] = _nearby(windows, cx, cy, thresh)

    return rooms


# ── Detection strategies ────────────────────────────────────────────────────

def _rooms_from_mahal_blocks(dxf_path: str, uf: float, min_area_m2: float) -> list[dict]:
    import ezdxf
    try:
        doc = ezdxf.readfile(dxf_path)
    except Exception:
        from ezdxf import recover
        doc, _ = recover.readfile(dxf_path)
    msp = doc.modelspace()

    rooms = []
    for ent in msp:
        if ent.dxftype() != "INSERT":
            continue
        layer = ent.dxf.layer.lower()
        if not any(kw in layer for kw in ("mahal", "room", "space", "0asm-mahal")):
            continue
        if not hasattr(ent, "attribs") or not ent.attribs:
            continue

        attrs = {a.dxf.tag.upper(): a.dxf.text for a in ent.attribs}
        name = (attrs.get("ROOMOBJECTS:NAME") or attrs.get("NAME") or
                attrs.get("MAHAL_ADI") or attrs.get("ROOM_NAME") or "")
        area_str = (attrs.get("ALAN:NAME") or attrs.get("ALAN") or
                    attrs.get("AREA") or "0")
        try:
            area_m2 = float(str(area_str).replace(",", ".").replace("m2", "").strip())
            if area_m2 > 1000:
                area_m2 *= uf
        except ValueError:
            area_m2 = 0.0

        if area_m2 < min_area_m2:
            continue

        cx, cy = ent.dxf.insert.x, ent.dxf.insert.y
        rooms.append({
            "id": len(rooms),
            "label": name.strip(),
            "area_m2": round(area_m2, 2),
            "centroid_x": cx,
            "centroid_y": cy,
            "points": [],
            "source": "mahal_block",
        })

    return rooms


def _rooms_from_hatch(dxf_path: str, uf: float, min_area_m2: float) -> list[dict]:
    import ezdxf
    try:
        doc = ezdxf.readfile(dxf_path)
    except Exception:
        from ezdxf import recover
        doc, _ = recover.readfile(dxf_path)
    msp = doc.modelspace()

    rooms = []
    for ent in msp:
        if ent.dxftype() != "HATCH":
            continue
        for path in ent.paths:
            pts = _boundary_path_to_points(path)
            if len(pts) < 3:
                continue
            try:
                from shapely.geometry import Polygon
                poly = Polygon(pts)
                if not poly.is_valid:
                    poly = poly.buffer(0)
                area_m2 = poly.area * uf
                if area_m2 < min_area_m2:
                    continue
                c = poly.centroid
                rooms.append({
                    "id": len(rooms),
                    "label": "",
                    "area_m2": round(area_m2, 2),
                    "centroid_x": c.x,
                    "centroid_y": c.y,
                    "points": list(poly.exterior.coords),
                    "source": "hatch",
                })
            except Exception:
                continue

    if len(rooms) >= 5:
        areas = sorted(r["area_m2"] for r in rooms)
        median = areas[len(areas) // 2]
        if median >= 3.0:
            return rooms
    return []


def _rooms_from_polygonize(entities: list, layer_types: dict, uf: float, min_area_m2: float) -> list[dict]:
    try:
        from shapely.ops import polygonize, unary_union
        from shapely.geometry import LineString, MultiLineString
    except ImportError:
        return []

    wall_lines = []
    for e in entities:
        lt = layer_types.get(e["layer"], "unknown")
        if lt not in ("walls", "columns"):
            continue
        if e["type"] == "LINE":
            wall_lines.append(LineString([(e["x1"], e["y1"]), (e["x2"], e["y2"])]))
        elif e["type"] == "LWPOLYLINE" and e.get("points"):
            pts = e["points"]
            for i in range(len(pts) - 1):
                wall_lines.append(LineString([pts[i], pts[i + 1]]))
            if e.get("closed") and len(pts) > 1:
                wall_lines.append(LineString([pts[-1], pts[0]]))

    if not wall_lines:
        return []

    merged = unary_union(wall_lines)
    polys = list(polygonize(merged))

    rooms = []
    for poly in polys:
        area_m2 = poly.area * uf
        if area_m2 < min_area_m2:
            continue
        c = poly.centroid
        rooms.append({
            "id": len(rooms),
            "label": "",
            "area_m2": round(area_m2, 2),
            "centroid_x": c.x,
            "centroid_y": c.y,
            "points": list(poly.exterior.coords),
            "source": "polygonize",
        })
    return rooms


def _rooms_from_closed_polylines(entities: list, layer_types: dict, uf: float, min_area_m2: float) -> list[dict]:
    try:
        from shapely.geometry import Polygon
    except ImportError:
        return []

    rooms = []
    for e in entities:
        if e["type"] != "LWPOLYLINE" or not e.get("closed"):
            continue
        pts = e.get("points", [])
        if len(pts) < 3:
            continue
        try:
            poly = Polygon(pts)
            area_m2 = poly.area * uf
            if area_m2 < min_area_m2:
                continue
            c = poly.centroid
            rooms.append({
                "id": len(rooms),
                "label": "",
                "area_m2": round(area_m2, 2),
                "centroid_x": c.x,
                "centroid_y": c.y,
                "points": list(poly.exterior.coords),
                "source": "closed_polyline",
            })
        except Exception:
            continue
    return rooms


# ── Helpers ──────────────────────────────────────────────────────────────────

def _boundary_path_to_points(path) -> list[tuple]:
    pts = []
    try:
        if hasattr(path, "vertices"):
            pts = [(v[0], v[1]) for v in path.vertices]
        elif hasattr(path, "edges"):
            for edge in path.edges:
                etype = edge.EDGE_TYPE if hasattr(edge, "EDGE_TYPE") else type(edge).__name__
                if "Line" in etype and hasattr(edge, "start"):
                    pts.append((edge.start.x, edge.start.y))
                elif "Arc" in etype and hasattr(edge, "center"):
                    import math as _m
                    cx, cy, r = edge.center.x, edge.center.y, edge.radius
                    a0, a1 = _m.radians(edge.start_angle), _m.radians(edge.end_angle)
                    if a1 < a0:
                        a1 += 2 * _m.pi
                    for i in range(8):
                        a = a0 + (a1 - a0) * i / 7
                        pts.append((cx + r * _m.cos(a), cy + r * _m.sin(a)))
                elif hasattr(edge, "control_points"):
                    pts += [(p[0], p[1]) for p in edge.control_points]
    except Exception:
        pass
    return pts


def _collect_elements(entities: list, layer_types: dict, etype: str) -> list[dict]:
    result = []
    for e in entities:
        if layer_types.get(e["layer"]) != etype:
            continue
        x = e.get("x") or e.get("x1") or (e["points"][0][0] if e.get("points") else 0)
        y = e.get("y") or e.get("y1") or (e["points"][0][1] if e.get("points") else 0)
        result.append({"x": x, "y": y, "layer": e["layer"], "type": e["type"]})
    return result


def _nearby(elements: list, cx: float, cy: float, threshold: float) -> int:
    return sum(1 for e in elements if math.hypot(e["x"] - cx, e["y"] - cy) < threshold)
