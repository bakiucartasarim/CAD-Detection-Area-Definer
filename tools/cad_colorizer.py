"""
cad_colorizer.py
GstarCAD/AutoCAD'de tespit edilen odaları 3 renkle işaretler.

  Yeşil  (MAHAL-YESIL,   color 3) → İsim tanımlı alan
  Mavi   (MAHAL-MAVI,    color 5) → İsimsiz tanımlı alan
  Kırmızı(MAHAL-KIRMIZI, color 1) → Tanımsız alan (label var, polygon yok)
"""
from __future__ import annotations
import math
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from parsers.dxf_parser import parse_dxf
from parsers.element_classifier import classify_all_layers
from tools.ifc_exporter import _collect_labels, _polygon_area


LAYER_GREEN  = "MAHAL-YESIL"
LAYER_BLUE   = "MAHAL-MAVI"
LAYER_RED    = "MAHAL-KIRMIZI"
MAX_AREA_M2  = 200.0   # bina dış sınırı filtresi


def colorize_rooms(dxf_path: str) -> dict:
    """
    DXF'i analiz eder, GstarCAD'deki açık çizime 3 renkli hatch uygular.
    GstarCAD açık ve aynı DXF yüklü olmalıdır.

    Returns:
        {"green": N, "blue": N, "red": N}
    """
    import win32com.client
    import pythoncom

    # ── DXF Analiz ───────────────────────────────────────────────────────────
    data        = parse_dxf(dxf_path)
    layer_types = classify_all_layers(data["layers"])
    uf          = data["unit_factor"]
    ls          = math.sqrt(uf)
    wall_layers = {n for n, t in layer_types.items() if t == "walls"}

    rooms = _build_room_list(data, wall_layers, ls)
    labels = _collect_labels(data["entities"], ls)
    _match_labels(rooms, labels)

    unmatched_labels = [
        labels[i] for i in range(len(labels))
        if not any(r.get("_label_idx") == i for r in rooms)
    ]
    # ikinci geçiş: hangi label index'lerinin kullanıldığını izle
    used = set()
    for room in rooms:
        if room["name"]:
            # en yakın label'ı bul (aynı mantık)
            best_i = _find_best_label(room, labels, used)
            if best_i >= 0:
                used.add(best_i)
    unmatched_labels = [labels[i] for i in range(len(labels)) if i not in used]

    # ── GstarCAD Bağlantısı ──────────────────────────────────────────────────
    import time as _time
    for attempt in range(3):
        try:
            acad = win32com.client.GetActiveObject("GstarCAD.Application")
            doc  = acad.ActiveDocument
            msp  = doc.ModelSpace
            # Bağlantı testi
            _ = msp.Count
            break
        except Exception:
            if attempt == 2:
                raise
            _time.sleep(3)

    _clear_layers(msp, [LAYER_GREEN, LAYER_BLUE, LAYER_RED,
                        "DUVAR-HATCH", "MAHAL-TANIMLI", "MAHAL-TANIMSIZ"])
    _ensure_layers(doc, {LAYER_GREEN: 3, LAYER_BLUE: 5, LAYER_RED: 1})

    green = blue = red = 0

    # Geçici polyline oluştur → hatch uygula → polyline sil
    # Bu yaklaşım mevcut nesneleri taramadan çalışır.
    for room in rooms:
        pts = room.get("points", [])
        if len(pts) < 3:
            continue
        layer = LAYER_GREEN if room["name"] else LAYER_BLUE
        color = 3           if room["name"] else 5
        try:
            # Koordinatları DXF biriminden CAD birimine çevir (ls tersine)
            flat = []
            for p in pts:
                flat.extend([p[0] / ls, p[1] / ls])
            coords_var = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8, flat)
            # Geçici kapalı polyline oluştur
            tmp_poly = msp.AddLightWeightPolyline(coords_var)
            tmp_poly.Closed = True
            tmp_poly.Layer = layer

            # Hatch oluştur ve geçici polyline'ı boundary olarak ekle
            outer = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [tmp_poly])
            h = msp.AddHatch(0, "SOLID", True)
            h.Layer = layer
            h.Color = color
            h.AppendOuterLoop(outer)
            h.Evaluate()

            # Geçici polyline'ı sil (hatch bağımsız kalır)
            tmp_poly.Delete()

            if room["name"]:
                green += 1
            else:
                blue += 1
        except Exception:
            pass

    # Kırmızı: daire + metin marker
    for lbl in unmatched_labels:
        dx = lbl["x"] / ls
        dy = lbl["y"] / ls
        r  = 500.0
        try:
            pt = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8, [dx, dy, 0.0])
            circ = msp.AddCircle(pt, r)
            circ.Layer = LAYER_RED
            circ.Color = 1

            tp = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8, [dx, dy + r * 0.3, 0.0])
            txt = msp.AddText(lbl["name"] + " (TANIMSIZ)", tp, r * 0.4)
            txt.Layer = LAYER_RED
            txt.Color = 1
            red += 1
        except Exception:
            pass

    doc.Regen(1)
    return {"green": green, "blue": blue, "red": red}


# ── Helpers ───────────────────────────────────────────────────────────────────

def _build_room_list(data, wall_layers, ls):
    rooms = []
    for ent in data["entities"]:
        if ent["type"] != "LWPOLYLINE":
            continue
        if ent["layer"] not in wall_layers:
            continue
        if not ent.get("closed"):
            continue
        pts = ent.get("points", [])
        if len(pts) < 3:
            continue
        pts_m = [[p[0] * ls, p[1] * ls] for p in pts]
        area  = _polygon_area(pts_m)
        if area > MAX_AREA_M2:
            continue
        cx = sum(p[0] for p in pts_m) / len(pts_m)
        cy = sum(p[1] for p in pts_m) / len(pts_m)
        rooms.append({"cx": cx, "cy": cy, "name": "", "area": area, "area_m2": area, "points": pts_m})
    return rooms


def _deduplicate_rooms(rooms: list) -> list:
    """Merkezi 1.5m içinde olan duplike polyline'ları at, en büyük olanı tut."""
    kept = []
    used_indices: set[int] = set()
    sorted_rooms = sorted(enumerate(rooms),
                          key=lambda x: x[1]["area"], reverse=True)
    for orig_idx, room in sorted_rooms:
        if orig_idx in used_indices:
            continue
        for other_idx, other in enumerate(rooms):
            if other_idx == orig_idx or other_idx in used_indices:
                continue
            if math.hypot(other["cx"] - room["cx"],
                          other["cy"] - room["cy"]) < 1.5:
                used_indices.add(other_idx)
        used_indices.add(orig_idx)
        kept.append(room)
    kept_set = set(id(r) for r in kept)
    return [r for r in rooms if id(r) in kept_set]


def _match_labels(rooms, labels):
    used: set[int] = set()
    for room in rooms:
        radius = math.sqrt(room["area"] / math.pi)
        thresh = max(radius * 3.0 + 1.0, 6.0)  # +1m buffer
        best_i = _find_best_label(room, labels, used)
        if best_i >= 0:
            d = math.hypot(labels[best_i]["x"] - room["cx"],
                           labels[best_i]["y"] - room["cy"])
            if d <= thresh:
                room["name"]   = labels[best_i]["name"]
                room["number"] = labels[best_i]["number"]
                used.add(best_i)


def _find_best_label(room, labels, used):
    best_i, best_d = -1, float("inf")
    for i, lbl in enumerate(labels):
        if i in used:
            continue
        d = math.hypot(lbl["x"] - room["cx"], lbl["y"] - room["cy"])
        if d < best_d:
            best_d, best_i = d, i
    return best_i


def _clear_layers(msp, layer_names):
    """Hedef layerlardaki nesneleri LISP komutuyla siler."""
    doc = msp.Application.ActiveDocument
    for layer_name in layer_names:
        try:
            cmd = f'(command "._ERASE" (ssget "X" (list (cons 8 "{layer_name}"))) "")\n'
            doc.SendCommand(cmd)
        except Exception:
            pass


def _ensure_layers(doc, layer_colors: dict):
    for name, color in layer_colors.items():
        try:
            l = doc.Layers.Add(name)
        except Exception:
            l = doc.Layers.Item(name)
        l.Color = color


def _collect_gcad_polys(msp, wall_layers, ls):
    """SelectionSet ile layer+tip filtresi uygular — tüm nesneleri taramaz."""
    import win32com.client, pythoncom
    doc = msp.Application.ActiveDocument
    polys = []

    for layer_name in wall_layers:
        sel_name = "_TMP_WALL_SEL"
        try:
            doc.SelectionSets.Item(sel_name).Delete()
        except Exception:
            pass
        try:
            sel = doc.SelectionSets.Add(sel_name)
            ft = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_I2,
                [0, 8])  # 0=DXF tipi, 8=layer
            fd = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_BSTR,
                ["LWPOLYLINE", layer_name])
            sel.Select(5, None, None, ft, fd)  # 5=ssget ALL
            for i in range(sel.Count):
                try:
                    e = sel.Item(i)
                    if not e.Closed:
                        continue
                    coords = e.Coordinates
                    n = len(coords) // 2
                    pts_m = [[coords[j * 2] * ls, coords[j * 2 + 1] * ls]
                             for j in range(n)]
                    if _polygon_area(pts_m) <= MAX_AREA_M2:
                        polys.append(e)
                except Exception:
                    pass
            sel.Delete()
        except Exception:
            pass

    return polys
