"""
dxf_parser.py
Raw entity extraction from DXF files using ezdxf.
Returns structured data: layers, entities, bounding box, unit factor.
"""
from __future__ import annotations
import math
from typing import Any


def parse_dxf(dxf_path: str) -> dict:
    """
    Parse a DXF file and return all entities grouped by layer and type.
    Returns:
        {
          "layers": {layer_name: {"types": [...], "count": N}},
          "entities": [...],
          "bbox": {"min_x", "min_y", "max_x", "max_y"},
          "unit_factor": float,   # 1e-6=mm, 1e-4=cm, 1.0=m
          "unit_label": str,
          "block_names": [...],
          "hatch_layers": [...],
        }
    """
    import ezdxf
    try:
        doc = ezdxf.readfile(dxf_path)
    except Exception:
        from ezdxf import recover
        doc, _ = recover.readfile(dxf_path)

    msp = doc.modelspace()
    layers: dict[str, dict] = {}
    entities: list[dict] = []
    all_x: list[float] = []
    all_y: list[float] = []
    block_names: set[str] = set()
    hatch_layers: set[str] = set()

    for ent in msp:
        layer = ent.dxf.layer
        etype = ent.dxftype()

        if layer not in layers:
            layers[layer] = {"types": set(), "count": 0}
        layers[layer]["types"].add(etype)
        layers[layer]["count"] += 1

        rec: dict[str, Any] = {"layer": layer, "type": etype}

        if etype == "LINE":
            sx, sy = ent.dxf.start.x, ent.dxf.start.y
            ex, ey = ent.dxf.end.x, ent.dxf.end.y
            rec.update({"x1": sx, "y1": sy, "x2": ex, "y2": ey,
                        "length": math.hypot(ex - sx, ey - sy)})
            all_x += [sx, ex]; all_y += [sy, ey]

        elif etype == "LWPOLYLINE":
            pts = [(p[0], p[1]) for p in ent.get_points()]
            rec.update({"points": pts, "closed": ent.closed, "vertex_count": len(pts)})
            all_x += [p[0] for p in pts]; all_y += [p[1] for p in pts]

        elif etype == "CIRCLE":
            cx, cy, r = ent.dxf.center.x, ent.dxf.center.y, ent.dxf.radius
            rec.update({"cx": cx, "cy": cy, "radius": r})
            all_x += [cx - r, cx + r]; all_y += [cy - r, cy + r]

        elif etype == "ARC":
            cx, cy = ent.dxf.center.x, ent.dxf.center.y
            rec.update({"cx": cx, "cy": cy, "radius": ent.dxf.radius,
                        "start_angle": ent.dxf.start_angle,
                        "end_angle": ent.dxf.end_angle})
            all_x.append(cx); all_y.append(cy)

        elif etype == "INSERT":
            bx, by = ent.dxf.insert.x, ent.dxf.insert.y
            block_names.add(ent.dxf.name)
            attrs = {}
            if hasattr(ent, "attribs"):
                for a in ent.attribs:
                    attrs[a.dxf.tag] = a.dxf.text
            rec.update({"x": bx, "y": by, "block": ent.dxf.name, "attrs": attrs})
            all_x.append(bx); all_y.append(by)

        elif etype == "HATCH":
            hatch_layers.add(layer)
            rec.update({"pattern": getattr(ent.dxf, "pattern_name", ""),
                        "path_count": len(ent.paths)})

        elif etype in ("TEXT", "MTEXT"):
            txt = ent.dxf.text if etype == "TEXT" else ent.text
            ix = ent.dxf.insert.x if hasattr(ent.dxf, "insert") else 0
            iy = ent.dxf.insert.y if hasattr(ent.dxf, "insert") else 0
            rec.update({"text": txt, "x": ix, "y": iy})
            all_x.append(ix); all_y.append(iy)

        entities.append(rec)

    # Serialise layer type sets → lists
    for ln in layers:
        layers[ln]["types"] = list(layers[ln]["types"])

    bbox = {}
    if all_x:
        bbox = {"min_x": min(all_x), "max_x": max(all_x),
                "min_y": min(all_y), "max_y": max(all_y)}

    unit_factor, unit_label = _detect_unit(bbox)

    return {
        "layers": layers,
        "entities": entities,
        "bbox": bbox,
        "unit_factor": unit_factor,
        "unit_label": unit_label,
        "block_names": sorted(block_names),
        "hatch_layers": sorted(hatch_layers),
        "entity_count": len(entities),
        "layer_count": len(layers),
    }


def _detect_unit(bbox: dict) -> tuple[float, str]:
    if not bbox:
        return 1e-4, "cm"
    dx = bbox["max_x"] - bbox["min_x"]
    dy = bbox["max_y"] - bbox["min_y"]
    for uf, label in [(1e-6, "mm"), (1e-4, "cm"), (1.0, "m")]:
        footprint = dx * dy * uf
        if 30 <= footprint <= 500_000:
            return uf, label
    return 1e-4, "cm"
