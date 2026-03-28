"""
parsers/room_matcher.py
=======================
Saf Python — MCP, COM, tkinter bağımlılığı yok.
Lumina ve server.py ortak olarak bu modülü kullanır.

Ana fonksiyon:
    match_rooms(dxf_path) → list[RoomMatch]

RoomMatch:
    id, name, number, area_m2, cx, cy, points, status, layer
    status: 'named' | 'unnamed' | 'unmatched'
"""
from __future__ import annotations
import math
import os
import sys
from dataclasses import dataclass, field
from typing import Optional

# parsers/ klasörünün parent'ını (proje kökü) path'e ekle
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)


@dataclass
class RoomMatch:
    id: int
    name: str
    number: str
    area_m2: float
    cx: float
    cy: float
    points: list = field(default_factory=list)   # [[x, y], ...]  ham DXF birimleri
    status: str = "unnamed"   # 'named' | 'unnamed' | 'unmatched'
    layer: str = ""


# ── Yardımcı geometri ──────────────────────────────────────────────────────

def _polygon_area(pts: list) -> float:
    """Shoelace (Gauss) formülü — pozitif alan döner."""
    n = len(pts)
    if n < 3:
        return 0.0
    a = 0.0
    for i in range(n):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % n]
        a += x1 * y2 - x2 * y1
    return abs(a) / 2.0


def _pt_in_poly(px: float, py: float, pts: list) -> bool:
    inside, j = False, len(pts) - 1
    for i in range(len(pts)):
        xi, yi = pts[i]
        xj, yj = pts[j]
        if ((yi > py) != (yj > py)) and (px < (xj - xi) * (py - yi) / (yj - yi + 1e-12) + xi):
            inside = not inside
        j = i
    return inside


def _is_closed_poly(ent: dict, gap_threshold: float = 50.0) -> bool:
    if ent.get("closed"):
        return True
    pts = ent.get("points", [])
    if len(pts) >= 3:
        gap = math.hypot(pts[-1][0] - pts[0][0], pts[-1][1] - pts[0][1])
        if gap < gap_threshold:
            return True
    return False


def _parse_area(raw: str) -> float:
    try:
        return float(raw.replace(",", ".").replace("m2", "").replace("m²", "").strip())
    except Exception:
        return 0.0


# ── Ana fonksiyon ──────────────────────────────────────────────────────────

def match_rooms(
    dxf_path: str,
    max_area_m2: float = 200.0,
    ai_mahal_layer: str = "AI_MAHAL",
) -> list[RoomMatch]:
    """
    DXF'teki MAHAL bloklarını kapalı LWPOLYLINE'larla eşleştirir.

    Adımlar:
      1. Duvar layer'larındaki kapalı polygon'ları topla
      2. AI_MAHAL layer'ındaki manuel polygon'ları ekle
      3. Merkezi yakın duplicate'leri kaldır (en büyüğü tut)
      4. MAHAL INSERT bloklarından label oku
      5. Greedy eşleştir: point-in-polygon → mesafe fallback

    Returns:
        list[RoomMatch]  — ham DXF birimleri (mm veya cm — INSUNITS'e göre)
    """
    import ezdxf as _ezdxf
    from parsers.dxf_parser import parse_dxf
    from parsers.element_classifier import classify_all_layers

    # ── 1. DXF oku ─────────────────────────────────────────────────────────
    try:
        doc = _ezdxf.readfile(dxf_path, encoding="utf-8")
    except Exception:
        doc = _ezdxf.readfile(dxf_path)

    data = parse_dxf(dxf_path)
    layer_types = classify_all_layers(data["layers"])
    uf = data["unit_factor"]       # m²/birim²  (1e-6 for mm, 1e-4 for cm)
    ls = math.sqrt(uf)             # birim → m  (alan hesabı için)

    wall_layers = {n for n, t in layer_types.items() if t == "walls"}

    # ── 2. Kapalı polygon'ları topla ───────────────────────────────────────
    polygons = []
    for ent in data["entities"]:
        if ent["type"] != "LWPOLYLINE" or not _is_closed_poly(ent):
            continue
        layer = ent["layer"]
        if layer not in wall_layers and layer != ai_mahal_layer:
            continue
        pts = ent.get("points", [])
        if len(pts) < 3:
            continue
        area_m2 = _polygon_area([[p[0] * ls, p[1] * ls] for p in pts])
        max_a = 1000.0 if layer == ai_mahal_layer else max_area_m2
        if area_m2 < 0.5 or area_m2 > max_a:
            continue
        cx = sum(p[0] for p in pts) / len(pts)
        cy = sum(p[1] for p in pts) / len(pts)
        polygons.append({"pts": pts, "cx": cx, "cy": cy,
                         "area_m2": area_m2, "layer": layer})

    # ── 3. Duplicate kaldır — merkezi 10 birim'den yakın → en büyüğü tut ──
    dedup, used = [], set()
    for i, p in enumerate(polygons):
        if i in used:
            continue
        group = [i]
        for j, q in enumerate(polygons):
            if j <= i or j in used:
                continue
            if math.hypot(p["cx"] - q["cx"], p["cy"] - q["cy"]) < 10:
                group.append(j)
        best = max(group, key=lambda k: polygons[k]["area_m2"])
        dedup.append(polygons[best])
        used.update(group)
    polygons = dedup

    # ── 4. MAHAL bloklarından label oku ────────────────────────────────────
    labels = []
    for e in doc.modelspace():
        if e.dxftype() != "INSERT":
            continue
        layer = e.dxf.layer.upper()
        bname = e.dxf.name.upper()
        if not ("MAHAL" in layer or "MAHAL" in bname or "0ASM" in layer):
            continue
        if not hasattr(e, "attribs") or not e.attribs:
            continue
        attrs = {a.dxf.tag.upper(): a.dxf.text for a in e.attribs}
        name   = (attrs.get("ROOMOBJECTS:NAME") or attrs.get("NAME") or
                  attrs.get("MAHAL") or "").strip()
        number = (attrs.get("ROOMOBJECTS:NUMBER") or attrs.get("NUMBER") or
                  attrs.get("MAHALNO") or "").strip()
        area_r = attrs.get("ALAN:NAME") or attrs.get("ALAN") or attrs.get("AREA") or "0"
        labels.append({
            "name": name, "number": number,
            "area_m2": _parse_area(area_r),
            "x": e.dxf.insert.x, "y": e.dxf.insert.y,
        })

    # ── 5. Greedy eşleştir: polygon → label ────────────────────────────────
    rooms: list[RoomMatch] = []
    used_labels: set[int] = set()

    for poly in polygons:
        # Önce polygon içindeki label'lar
        inside = [i for i, lbl in enumerate(labels)
                  if i not in used_labels
                  and _pt_in_poly(lbl["x"], lbl["y"], poly["pts"])]
        if inside:
            best_i = min(inside, key=lambda i: math.hypot(
                labels[i]["x"] - poly["cx"], labels[i]["y"] - poly["cy"]))
        else:
            thresh = (math.sqrt(poly["area_m2"]) * 2.0 + 1.0) / ls
            cands  = [(i, math.hypot(lbl["x"] - poly["cx"], lbl["y"] - poly["cy"]))
                      for i, lbl in enumerate(labels) if i not in used_labels]
            cands  = [(i, d) for i, d in cands if d <= thresh]
            best_i = min(cands, key=lambda x: x[1])[0] if cands else -1

        if best_i >= 0:
            lbl = labels[best_i]
            used_labels.add(best_i)
            rooms.append(RoomMatch(
                id=len(rooms), name=lbl["name"], number=lbl["number"],
                area_m2=lbl["area_m2"],
                cx=round(poly["cx"], 1), cy=round(poly["cy"], 1),
                points=[[round(p[0], 1), round(p[1], 1)] for p in poly["pts"]],
                status="named" if (lbl["name"] or lbl["number"]) else "unnamed",
                layer=poly["layer"],
            ))
        else:
            rooms.append(RoomMatch(
                id=len(rooms), name="", number="",
                area_m2=round(poly["area_m2"], 2),
                cx=round(poly["cx"], 1), cy=round(poly["cy"], 1),
                points=[[round(p[0], 1), round(p[1], 1)] for p in poly["pts"]],
                status="unnamed", layer=poly["layer"],
            ))

    # Eşleşemeyen label'lar
    for i, lbl in enumerate(labels):
        if i not in used_labels and lbl["name"]:
            rooms.append(RoomMatch(
                id=len(rooms), name=lbl["name"], number=lbl["number"],
                area_m2=lbl["area_m2"],
                cx=round(lbl["x"], 1), cy=round(lbl["y"], 1),
                points=[], status="unmatched",
            ))

    return rooms


def match_rooms_json(dxf_path: str, **kwargs) -> dict:
    """match_rooms() sonucunu JSON-serializable dict olarak döner."""
    rooms = match_rooms(dxf_path, **kwargs)
    named     = sum(1 for r in rooms if r.status == "named")
    unnamed   = sum(1 for r in rooms if r.status == "unnamed")
    unmatched = sum(1 for r in rooms if r.status == "unmatched")
    return {
        "named": named, "unnamed": unnamed, "unmatched": unmatched,
        "toplam": len(rooms),
        "rooms": [
            {"id": r.id, "name": r.name, "number": r.number,
             "area_m2": r.area_m2, "cx": r.cx, "cy": r.cy,
             "points": r.points, "status": r.status, "layer": r.layer}
            for r in rooms
        ],
    }
