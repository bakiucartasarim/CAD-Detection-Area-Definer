"""
server.py — CAD Detection Area Definer MCP Server

Claude'a şu araçları açar:
  • analyze_cad              → DXF katmanları ve entity sayıları
  • detect_rooms             → Oda poligonları, alanları ve etiketleri
  • classify_elements        → Katmanları duvar/kapı/pencere/mobilya olarak sınıflandırır
  • get_unknown_layers       → Tanımlanamayan katmanları listeler
  • train_layer              → Yeni katman → tip eşleştirmesi öğretir
  • get_room_geometry        → Tek bir odanın tam geometrisi
  • export_walls_ifc         → Duvar polyline'larını IFC olarak dışa aktarır
  • colorize_rooms_in_cad    → GstarCAD'de polygon odaları renklendir (ADIM 0)
  • clean_lighting           → Armatür + hatch temizle (ADIM 1)
  • clean_cables             → Kablo/hat layer temizle (ADIM 2)
  • clean_block_hatches      → Blok içi hatch temizle (ADIM 3)
  • clean_hatch              → Modelspace hatch temizle (ADIM 4)
  • delete_tefris            → Tefris/mobilya temizle (ADIM 5)
  • delete_ceiling           → Tavan/asma tavan temizle (ADIM 6)
  • delete_linye             → Linye numaraları temizle (ADIM 7)
  • delete_electric_component→ Elektrik bileşen temizle (ADIM 8)
  • colorize_mahal_blocks    → MAHAL bloklarını GstarCAD'de renklendir (ADIM 9)
"""
from __future__ import annotations
import json
import sys
import os

# ── Korunan layer'lar — HİÇBİR adım bunları silmez ─────────────────────────
# Oda adı/numarası/alanı gibi mimari bilgi layer'larıdır.
_PROTECTED_LAYERS = {
    # Oda bilgi blokları (MAHAL NO, ROOMOBJECTS vb.)
    "0ASM-MAHAL BİLGİ", "0ASM-MAHAL BILGI", "0ASM-MAHAL",
    "_AB_MAHALMETIN", "_AB_MAHAL",
    "MAHAL_ISMI", "MAHAL_BILGI", "MAHAL BILGI",
    "MAHAL NO", "ROOMOBJECTS", "ROOM_INFO",
    # Duvar/yapı layer'ları
    "0ASM-DUVAR", "0ASM-KAPI", "0ASM-PENCERE",
    "0ASM-SIVA İÇ", "0ASM-SIVA IC", "0ASM-TRETUVAR",
    # Kolon/taşıyıcı sistem
    "KOL", "KOLON",
}

def _is_protected(layer_name: str) -> bool:
    """Bu layer hiçbir temizleme adımında silinmemeli."""
    lu = layer_name.upper()
    if layer_name in _PROTECTED_LAYERS or lu in {l.upper() for l in _PROTECTED_LAYERS}:
        return True
    # MAHAL içeren ama colorizer çıktısı olmayan layer'lar
    if "MAHAL" in lu and not any(x in lu for x in ("KIRMIZI", "YESIL", "MAVI", "RED", "GREEN", "BLUE")):
        return True
    return False

# Proje kökünü Python path'e ekle
sys.path.insert(0, os.path.dirname(__file__))

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("CAD Detection Area Definer")


@mcp.tool()
def analyze_cad(dxf_path: str) -> str:
    """
    DXF dosyasını parse eder. Katman listesi, entity sayıları,
    bounding box, birim (mm/cm/m) ve blok adlarını döner.

    Args:
        dxf_path: DXF dosyasının tam yolu
    """
    from parsers.dxf_parser import parse_dxf
    from parsers.element_classifier import classify_all_layers

    data = parse_dxf(dxf_path)
    layer_types = classify_all_layers(data["layers"])

    summary = {
        "file": os.path.basename(dxf_path),
        "entity_count": data["entity_count"],
        "layer_count": data["layer_count"],
        "unit": data["unit_label"],
        "unit_factor": data["unit_factor"],
        "bbox": data["bbox"],
        "block_names": data["block_names"][:20],
        "hatch_layers": data["hatch_layers"],
        "layers": {
            name: {
                "element_type": layer_types.get(name, "unknown"),
                "entity_types": info["types"],
                "count": info["count"],
            }
            for name, info in data["layers"].items()
        },
    }
    return json.dumps(summary, ensure_ascii=False, indent=2)


@mcp.tool()
def detect_rooms(dxf_path: str, min_area_m2: float = 1.0) -> str:
    """
    DXF dosyasındaki odaları/mekanları tespit eder.
    Oda adı, alan (m²), merkez koordinatı, poligon noktaları,
    yakındaki kapı ve pencere sayısını döner.

    Tespit önceliği:
      1. MAHAL BLOCK'lardan (INSERT + attribs)
      2. HATCH boundary'lerden
      3. Duvar çizgilerini polygonize ederek
      4. Kapalı LWPOLYLINE'lardan

    Args:
        dxf_path: DXF dosyasının tam yolu
        min_area_m2: Minimum oda alanı (m²), varsayılan 1.0
    """
    from parsers.geometry_engine import detect_rooms as _detect

    rooms = _detect(dxf_path, min_area_m2=min_area_m2)

    result = {
        "total_rooms": len(rooms),
        "total_area_m2": round(sum(r["area_m2"] for r in rooms), 2),
        "detection_source": rooms[0]["source"] if rooms else "none",
        "rooms": [
            {
                "id": r["id"],
                "label": r["label"] or f"Oda {r['id'] + 1}",
                "area_m2": r["area_m2"],
                "centroid_x": round(r["centroid_x"], 2),
                "centroid_y": round(r["centroid_y"], 2),
                "has_polygon": len(r["points"]) > 0,
                "vertex_count": len(r["points"]),
                "doors_nearby": r.get("doors_nearby", 0),
                "windows_nearby": r.get("windows_nearby", 0),
                "source": r["source"],
            }
            for r in rooms
        ],
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def get_room_geometry(dxf_path: str, room_id: int, min_area_m2: float = 1.0) -> str:
    """
    Belirli bir odanın tam geometrisini (poligon noktaları) döner.

    Args:
        dxf_path: DXF dosyasının tam yolu
        room_id: detect_rooms'dan gelen oda ID'si (0 tabanlı)
        min_area_m2: Minimum oda alanı (m²)
    """
    from parsers.geometry_engine import detect_rooms as _detect

    rooms = _detect(dxf_path, min_area_m2=min_area_m2)
    room = next((r for r in rooms if r["id"] == room_id), None)

    if room is None:
        return json.dumps({"error": f"Room ID {room_id} not found. Total: {len(rooms)}"})

    return json.dumps({
        "id": room["id"],
        "label": room["label"] or f"Oda {room['id'] + 1}",
        "area_m2": room["area_m2"],
        "centroid_x": round(room["centroid_x"], 2),
        "centroid_y": round(room["centroid_y"], 2),
        "points": [[round(p[0], 2), round(p[1], 2)] for p in room["points"]],
        "source": room["source"],
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def classify_elements(dxf_path: str) -> str:
    """
    DXF dosyasındaki tüm katmanları semantik tiplere sınıflandırır:
    duvar, kapı, pencere, mobilya, kolon, merdiven, oda, metin, elektrik, bilinmiyor.

    Args:
        dxf_path: DXF dosyasının tam yolu
    """
    from parsers.dxf_parser import parse_dxf
    from parsers.element_classifier import classify_all_layers

    data = parse_dxf(dxf_path)
    layer_types = classify_all_layers(data["layers"])

    # Tiplere göre grupla
    grouped: dict[str, list] = {}
    for layer, etype in layer_types.items():
        grouped.setdefault(etype, []).append({
            "layer": layer,
            "entity_types": data["layers"][layer]["types"],
            "count": data["layers"][layer]["count"],
        })

    summary = {
        "file": os.path.basename(dxf_path),
        "classified": {k: v for k, v in grouped.items() if k != "unknown"},
        "unknown_layers": grouped.get("unknown", []),
        "unknown_count": len(grouped.get("unknown", [])),
    }
    return json.dumps(summary, ensure_ascii=False, indent=2)


@mcp.tool()
def get_unknown_layers(dxf_path: str) -> str:
    """
    Sistemin tanıyamadığı (sınıflandıramadığı) katman adlarını listeler.
    Bu katmanları train_layer() ile öğretebilirsiniz.

    Args:
        dxf_path: DXF dosyasının tam yolu
    """
    from parsers.dxf_parser import parse_dxf
    from parsers.element_classifier import get_unknown_layers as _get_unknown

    data = parse_dxf(dxf_path)
    unknown = _get_unknown(data["layers"])

    return json.dumps({
        "file": os.path.basename(dxf_path),
        "unknown_layers": unknown,
        "count": len(unknown),
        "tip": "Use train_layer(layer_name, element_type) to teach these layers.",
        "valid_types": [
            "walls", "doors", "windows", "furniture",
            "columns", "stairs", "rooms", "dimensions",
            "text", "electrical"
        ],
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def train_layer(layer_name: str, element_type: str) -> str:
    """
    Sisteme yeni bir katman → eleman tipi eşleştirmesi öğretir.
    Kalıcı olarak layer_registry.json'a kaydedilir.

    Args:
        layer_name: DXF katman adı (örn: "DUVARLAR", "0ASM-KAPI")
        element_type: Eleman tipi: walls | doors | windows | furniture |
                      columns | stairs | rooms | dimensions | text | electrical
    """
    from parsers.element_classifier import train_layer as _train

    result = _train(layer_name, element_type)
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def export_walls_ifc(
    dxf_path: str,
    output_path: str = "",
    wall_height_m: float = 3.0,
    wall_thickness_m: float = 0.10,
) -> str:
    """
    DXF'teki kapalı duvar polyline'larını IFC dosyasına dışa aktarır.
    Her kapalı alan → IfcSpace (oda) + IfcWall (10cm kalınlık, 3m yükseklik).

    Args:
        dxf_path: DXF dosyasının tam yolu
        output_path: Çıktı IFC dosyası yolu (boş bırakılırsa DXF ile aynı klasöre kaydeder)
        wall_height_m: Duvar yüksekliği metre cinsinden (varsayılan: 3.0)
        wall_thickness_m: Duvar kalınlığı metre cinsinden (varsayılan: 0.10)
    """
    from tools.ifc_exporter import export_walls_to_ifc

    if not output_path:
        base = os.path.splitext(dxf_path)[0]
        output_path = base + "_walls.ifc"

    result = export_walls_to_ifc(
        dxf_path=dxf_path,
        output_path=output_path,
        wall_height_m=wall_height_m,
        wall_thickness_m=wall_thickness_m,
    )
    return json.dumps(result, ensure_ascii=False, indent=2)


# ── İŞARETLEME ADIMLARI ──────────────────────────────────────────────────────

@mcp.tool()
def identify_rooms(dxf_path: str) -> str:
    """
    İŞARETLEME ADIM 1 — Mekan isimlerini belirle.
    DXF'teki MAHAL bloklarından oda adı, numarası ve alanını okur.
    colorize_rooms_in_cad() öncesinde çalıştırılır.

    Args:
        dxf_path: Temizlenmiş DXF dosyasının tam yolu
    Returns:
        JSON: rooms listesi [{id, name, number, area_m2, x, y}]
    """
    import ezdxf as _ezdxf

    def _parse_area(raw: str) -> float:
        try:
            return float(raw.replace(",", ".").replace("m2", "").replace("m²", "").strip())
        except Exception:
            return 0.0

    # utf-8 ile oku — GstarCAD header'da ANSI_1254 yazar ama içerik UTF-8'dir
    try:
        doc = _ezdxf.readfile(dxf_path, encoding="utf-8")
    except Exception:
        doc = _ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    rooms = []
    for e in msp:
        if e.dxftype() != "INSERT":
            continue
        layer = e.dxf.layer.upper()
        bname = e.dxf.name.upper()
        if not ("MAHAL" in layer or "MAHAL" in bname or "ROOM" in layer or "0ASM" in layer):
            continue
        if not hasattr(e, "attribs") or not e.attribs:
            continue

        attrs = {a.dxf.tag.upper(): a.dxf.text for a in e.attribs}

        name = (attrs.get("ROOMOBJECTS:NAME") or attrs.get("NAME") or
                attrs.get("MAHAL_ADI") or attrs.get("MAHAL") or "")
        number = (attrs.get("ROOMOBJECTS:NUMBER") or attrs.get("NUMBER") or
                  attrs.get("MAHALNO") or "")
        area_raw = (attrs.get("ALAN:NAME") or attrs.get("ALAN") or
                    attrs.get("AREA") or "0")

        name   = name.strip()
        number = number.strip()
        area_m2 = _parse_area(area_raw)

        if not name and not number:
            continue

        rooms.append({
            "id": len(rooms),
            "name": name,
            "number": number,
            "area_m2": area_m2,
            "x": round(e.dxf.insert.x, 1),
            "y": round(e.dxf.insert.y, 1),
            "layer": e.dxf.layer,
        })

    # numaraya göre sırala
    rooms.sort(key=lambda r: r["number"])
    for i, r in enumerate(rooms):
        r["id"] = i

    result = {
        "toplam": len(rooms),
        "rooms": rooms,
        "sonraki_adim": "match_rooms_to_polygons() ile polygon eşleştirmesi yapın"
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def match_rooms_to_polygons(dxf_path: str) -> str:
    """
    İŞARETLEME ADIM 2 — Her MAHAL bloğunu en yakın kapalı polygon'a eşleştir.
    identify_rooms() sonrasında çalıştırılır.

    Duvar layer'larındaki kapalı LWPOLYLINE'ları tarar, her birini
    içinde ya da en yakınında olan MAHAL bloğuyla eşleştirir.

    Args:
        dxf_path: Temizlenmiş DXF dosyasının tam yolu
    Returns:
        JSON: matched_rooms listesi [{id, name, number, area_m2, points, cx, cy, status}]
        status: 'named' | 'unnamed' | 'unmatched'
    """
    import ezdxf as _ezdxf
    import math as _math

    from parsers.dxf_parser import parse_dxf
    from parsers.element_classifier import classify_all_layers
    from tools.ifc_exporter import _polygon_area

    MAX_AREA_M2 = 200.0

    # ── 1. DXF oku — ham koordinatlar (ölçeksiz, GstarCAD birimleri) ─────────
    try:
        doc_enc = _ezdxf.readfile(dxf_path, encoding="utf-8")
    except Exception:
        doc_enc = _ezdxf.readfile(dxf_path)

    data        = parse_dxf(dxf_path)
    layer_types = classify_all_layers(data["layers"])
    uf          = data["unit_factor"]   # m²/birim²  (ör: 1e-6 for mm)
    ls          = _math.sqrt(uf)        # sadece alan hesabı için
    wall_layers = {n for n, t in layer_types.items() if t == "walls"}
    # AI_MAHAL: kullanıcının manuel çizdiği + algoritmik polygon'lar
    AI_MAHAL_LAYER = "AI_MAHAL"

    # ── 2. Kapalı polygon'ları topla (ham DXF birimleri) ───────────────────
    def _is_closed_poly(ent):
        if ent.get("closed"):
            return True
        # Geometrik kapalılık: ilk ve son nokta 50mm'den yakınsa kapalı say
        pts = ent.get("points", [])
        if len(pts) >= 3:
            import math as _m
            gap = _m.hypot(pts[-1][0]-pts[0][0], pts[-1][1]-pts[0][1])
            if gap < 50:
                return True
        return False

    polygons = []
    for ent in data["entities"]:
        if ent["type"] != "LWPOLYLINE" or not _is_closed_poly(ent):
            continue
        layer = ent["layer"]
        # Duvar layer'ları VEYA AI_MAHAL (manuel çizimler)
        if layer not in wall_layers and layer != AI_MAHAL_LAYER:
            continue
        pts = ent.get("points", [])
        if len(pts) < 3:
            continue
        area_m2 = _polygon_area([[p[0]*ls, p[1]*ls] for p in pts])
        # AI_MAHAL polygon'ları için alan filtresini genişlet (HOL gibi büyük alanlar)
        max_area = 1000.0 if layer == AI_MAHAL_LAYER else MAX_AREA_M2
        if area_m2 < 0.5 or area_m2 > max_area:
            continue
        cx = sum(p[0] for p in pts) / len(pts)
        cy = sum(p[1] for p in pts) / len(pts)
        polygons.append({"pts": pts, "cx": cx, "cy": cy, "area_m2": area_m2,
                         "layer": layer})

    # ── 3. MAHAL blok bilgilerini oku (ham DXF birimleri) ──────────────────
    def _parse_area(raw: str) -> float:
        try:
            return float(raw.replace(",", ".").replace("m2", "").replace("m²", "").strip())
        except Exception:
            return 0.0

    labels = []
    for e in doc_enc.modelspace():
        if e.dxftype() != "INSERT":
            continue
        layer = e.dxf.layer.upper()
        bname = e.dxf.name.upper()
        if not ("MAHAL" in layer or "MAHAL" in bname or "0ASM" in layer):
            continue
        if not hasattr(e, "attribs") or not e.attribs:
            continue
        attrs = {a.dxf.tag.upper(): a.dxf.text for a in e.attribs}
        name     = (attrs.get("ROOMOBJECTS:NAME") or attrs.get("NAME") or attrs.get("MAHAL") or "").strip()
        number   = (attrs.get("ROOMOBJECTS:NUMBER") or attrs.get("NUMBER") or attrs.get("MAHALNO") or "").strip()
        area_raw = attrs.get("ALAN:NAME") or attrs.get("ALAN") or attrs.get("AREA") or "0"
        lx = e.dxf.insert.x
        ly = e.dxf.insert.y
        labels.append({"name": name, "number": number, "area_m2": _parse_area(area_raw), "x": lx, "y": ly})

    # Deduplication: aynı odanın SIVAA+duvar+0ASM-DUVAR polygon'larını birleştir
    # Merkezi 10mm'den yakın olanlar aynı odanın farklı layer versiyonu — en büyüğünü tut
    dedup = []
    used_idx = set()
    for i, p in enumerate(polygons):
        if i in used_idx:
            continue
        group = [i]
        for j, q in enumerate(polygons):
            if j <= i or j in used_idx:
                continue
            if _math.hypot(p["cx"] - q["cx"], p["cy"] - q["cy"]) < 10:
                group.append(j)
        best = max(group, key=lambda k: polygons[k]["area_m2"])
        dedup.append(polygons[best])
        used_idx.update(group)
    polygons = dedup

    # ── 4. Eşleştir: polygon→label (greedy, en yakın label) ──────────────
    def _pt_in_poly(px, py, pts):
        inside = False
        n = len(pts)
        j = n - 1
        for i in range(n):
            xi, yi = pts[i]
            xj, yj = pts[j]
            if ((yi > py) != (yj > py)) and (px < (xj-xi)*(py-yi)/(yj-yi+1e-12) + xi):
                inside = not inside
            j = i
        return inside

    matched_rooms = []
    used_labels   = set()

    for poly in polygons:
        # Önce polygon içindeki label'ları dene
        inside = [i for i, lbl in enumerate(labels)
                  if i not in used_labels and _pt_in_poly(lbl["x"], lbl["y"], poly["pts"])]

        if inside:
            best_i = min(inside, key=lambda i: _math.hypot(
                labels[i]["x"] - poly["cx"], labels[i]["y"] - poly["cy"]))
        else:
            thresh = (_math.sqrt(poly["area_m2"]) * 2.0 + 1.0) / ls
            cands  = [(i, _math.hypot(lbl["x"]-poly["cx"], lbl["y"]-poly["cy"]))
                      for i, lbl in enumerate(labels) if i not in used_labels]
            cands  = [(i, d) for i, d in cands if d <= thresh]
            best_i = min(cands, key=lambda x: x[1])[0] if cands else -1

        if best_i >= 0:
            lbl = labels[best_i]
            used_labels.add(best_i)
            matched_rooms.append({
                "id":      len(matched_rooms),
                "name":    lbl["name"],
                "number":  lbl["number"],
                "area_m2": lbl["area_m2"],
                "cx":      round(poly["cx"], 1),
                "cy":      round(poly["cy"], 1),
                "points":  [[round(p[0], 1), round(p[1], 1)] for p in poly["pts"]],
                "status":  "named" if (lbl["name"] or lbl["number"]) else "unnamed",
            })
        else:
            matched_rooms.append({
                "id":      len(matched_rooms),
                "name":    "",
                "number":  "",
                "area_m2": round(poly["area_m2"], 2),
                "cx":      round(poly["cx"], 1),
                "cy":      round(poly["cy"], 1),
                "points":  [[round(p[0], 1), round(p[1], 1)] for p in poly["pts"]],
                "status":  "unnamed",
            })

    # Eşleşemeyen label'lar (kırmızı)
    for i, lbl in enumerate(labels):
        if i not in used_labels and lbl["name"]:
            matched_rooms.append({
                "id":      len(matched_rooms),
                "name":    lbl["name"],
                "number":  lbl["number"],
                "area_m2": lbl["area_m2"],
                "cx":      round(lbl["x"], 1),
                "cy":      round(lbl["y"], 1),
                "points":  [],
                "status":  "unmatched",
            })

    named    = sum(1 for r in matched_rooms if r["status"] == "named")
    unnamed  = sum(1 for r in matched_rooms if r["status"] == "unnamed")
    unmatched = sum(1 for r in matched_rooms if r["status"] == "unmatched")

    result = {
        "named":     named,
        "unnamed":   unnamed,
        "unmatched": unmatched,
        "toplam":    len(matched_rooms),
        "rooms":     matched_rooms,
        "sonraki_adim": "draw_room_markers() ile GstarCAD'e hatch + etiket çizin",
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def request_manual_polygons(dxf_path: str) -> str:
    """
    İŞARETLEME ADIM 3.5 — Tanımlanamayan mekanlar için kullanıcıdan manuel polygon ister.
    match_rooms_to_polygons() sonrasında, draw_room_markers() öncesinde çalıştırılır.

    1. Polygon bulunamayan odaları listeler
    2. GstarCAD'de AI_MAHAL layer'ını oluşturur ve aktif yapar
    3. Kullanıcıdan AI_MAHAL layer'ına kapalı LWPOLYLINE çizmesini ister

    Args:
        dxf_path: Temizlenmiş DXF dosyasının tam yolu
    """
    import win32com.client, json as _json
    import time as _time

    matched = _json.loads(match_rooms_to_polygons(dxf_path))
    unmatched = [r for r in matched["rooms"] if r["status"] == "unmatched"]

    if not unmatched:
        return _json.dumps({
            "mesaj": "Tüm mekanlar tespit edildi, manuel polygon gerekmiyor.",
            "unmatched": 0
        }, ensure_ascii=False, indent=2)

    # GstarCAD bağlantısı
    for attempt in range(3):
        try:
            acad = win32com.client.GetActiveObject("GstarCAD.Application")
            doc  = acad.ActiveDocument
            msp  = doc.ModelSpace
            _ = msp.Count
            break
        except Exception:
            if attempt == 2:
                raise
            _time.sleep(3)

    # AI_MAHAL layer'ını oluştur ve aktif yap
    AI_LAYER = "AI_MAHAL"
    try:
        ly = doc.Layers.Add(AI_LAYER)
    except Exception:
        ly = doc.Layers.Item(AI_LAYER)
    ly.Color = 4  # cyan
    doc.ActiveLayer = ly

    oda_listesi = [
        {"number": r["number"], "name": r["name"], "area_m2": r["area_m2"],
         "cx": r["cx"], "cy": r["cy"]}
        for r in unmatched
    ]

    result = {
        "unmatched_count": len(unmatched),
        "talimat": (
            f"GstarCAD'de {len(unmatched)} mekan kırmızı daire ile işaretlendi. "
            f"AI_MAHAL layer'ı aktif yapıldı (cyan). "
            f"Lütfen her kırmızı daireli mekan için AI_MAHAL layer'ına kapalı LWPOLYLINE çizin, "
            f"sonra draw_room_markers() çalıştırın."
        ),
        "eksik_odalar": oda_listesi,
        "sonraki_adim": "Polygon çizimi sonrası draw_room_markers() çalıştırın"
    }
    return _json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def draw_room_markers(dxf_path: str) -> str:
    """
    İŞARETLEME ADIM 3 — GstarCAD'de hatch + metin etiketi çiz.
    match_rooms_to_polygons() verilerini kullanarak aktif çizime uygular.

      Yeşil  (MAHAL-YESIL,   color 3) → İsim tanımlı alan + etiket
      Mavi   (MAHAL-MAVI,    color 5) → İsimsiz alan
      Kırmızı(MAHAL-KIRMIZI, color 1) → Tanımsız (polygon yok) → daire + metin

    Args:
        dxf_path: Temizlenmiş DXF dosyasının tam yolu
    """
    import win32com.client, pythoncom, math as _math, json as _json

    # ADIM 2 verisini al
    matched = _json.loads(match_rooms_to_polygons(dxf_path))
    rooms   = matched["rooms"]

    # GstarCAD bağlantısı
    import time as _time
    for attempt in range(3):
        try:
            acad = win32com.client.GetActiveObject("GstarCAD.Application")
            doc  = acad.ActiveDocument
            msp  = doc.ModelSpace
            _ = msp.Count
            break
        except Exception:
            if attempt == 2:
                raise
            _time.sleep(3)

    LAYER_GREEN  = "MAHAL-YESIL"
    LAYER_BLUE   = "MAHAL-MAVI"
    LAYER_RED    = "MAHAL-KIRMIZI"
    LABEL_LAYER  = "MAHAL-ETIKET"
    AI_MAHAL     = "AI_MAHAL"

    # Eski işaretleme layer'larını temizle (AI_MAHAL korunur — kullanıcı çizimleri)
    for ln in [LAYER_GREEN, LAYER_BLUE, LAYER_RED, LABEL_LAYER,
               "DUVAR-HATCH", "MAHAL-TANIMLI", "MAHAL-TANIMSIZ"]:
        try:
            doc.SendCommand(f'(command "._ERASE" (ssget "X" (list (cons 8 "{ln}"))) "")\n')
        except Exception:
            pass

    # Layer'ları oluştur — AI_MAHAL: kullanıcı manuel polygon layer'ı (cyan, 4)
    for lname, lcolor in [(LAYER_GREEN, 3), (LAYER_BLUE, 5), (LAYER_RED, 1),
                          (LABEL_LAYER, 7), (AI_MAHAL, 4)]:
        try:
            ly = doc.Layers.Add(lname)
        except Exception:
            ly = doc.Layers.Item(lname)
        ly.Color = lcolor

    green = blue = red = 0

    for room in rooms:
        pts    = room.get("points", [])
        status = room["status"]
        name   = room["name"]
        number = room["number"]
        area   = room["area_m2"]
        cx, cy = room["cx"], room["cy"]

        if status in ("named", "unnamed") and len(pts) >= 3:
            layer = LAYER_GREEN if status == "named" else LAYER_BLUE
            color = 3           if status == "named" else 5
            try:
                flat = []
                for p in pts:
                    flat.extend([p[0], p[1]])
                coords_var = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, flat)
                tmp_poly = msp.AddLightWeightPolyline(coords_var)
                tmp_poly.Closed = True
                tmp_poly.Layer  = layer

                outer = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [tmp_poly])
                h = msp.AddHatch(0, "SOLID", True)
                h.Layer = layer
                h.Color = color
                h.AppendOuterLoop(outer)
                h.Evaluate()
                tmp_poly.Delete()

                if status == "named":
                    green += 1
                else:
                    blue += 1
            except Exception:
                pass

            # Metin etiketi: no + isim + alan — polygon üst bölgesine yerleştir
            if status == "named" and name:
                try:
                    # Polygon bounding box (DXF birimleri = mm)
                    all_x = [p[0] for p in pts]
                    all_y = [p[1] for p in pts]
                    min_x, max_x = min(all_x), max(all_x)
                    min_y, max_y = min(all_y), max(all_y)
                    w_span = max_x - min_x
                    h_span = max_y - min_y
                    mid_x  = (min_x + max_x) / 2.0

                    # Metin yüksekliği: oda yüksekliğinin %8'i
                    # DXF birimleri mm → oda 4000mm ise txt_h=320mm (uygun)
                    txt_h = max(min(h_span * 0.006, 75), 15)

                    # Üst bölge: max_y'den %20 aşağı
                    ty = max_y - h_span * 0.20

                    # MTEXT — sol kenar = mid_x - w_span/2, genişlik = w_span
                    insert1 = win32com.client.VARIANT(
                        pythoncom.VT_ARRAY | pythoncom.VT_R8, [min_x, ty, 0.0])
                    mt1 = msp.AddMText(insert1, w_span,
                                       f"\\A1;{number}  {name}" if number else f"\\A1;{name}")
                    mt1.Layer  = LABEL_LAYER
                    mt1.Color  = 7
                    mt1.Height = txt_h

                    insert2 = win32com.client.VARIANT(
                        pythoncom.VT_ARRAY | pythoncom.VT_R8, [min_x, ty - txt_h * 1.6, 0.0])
                    mt2 = msp.AddMText(insert2, w_span, f"\\A1;{area} m\u00b2")
                    mt2.Layer  = LABEL_LAYER
                    mt2.Color  = 7
                    mt2.Height = txt_h * 0.75
                except Exception:
                    pass

        elif status == "unmatched":
            # Polygon yok — MAHAL bloğu merkezinde küçük kırmızı daire
            try:
                r_circ = 150.0
                pt = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [cx, cy, 0.0])
                circ = msp.AddCircle(pt, r_circ)
                circ.Layer = LAYER_RED
                circ.Color = 1
                red += 1
            except Exception:
                pass

    doc.Regen(1)

    result = {
        "yesil":     green,
        "mavi":      blue,
        "kirmizi":   red,
        "summary":   f"{green} yeşil (isim+etiket), {blue} mavi (isimsiz), {red} kırmızı (tanımsız)",
        "sonraki_adim": "export_room_report() ile mekan listesini dışa aktarın",
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def colorize_rooms_in_cad(dxf_path: str) -> str:
    """
    GstarCAD/AutoCAD'de açık olan çizime 3 renkli hatch uygular.
    GstarCAD açık ve dxf_path ile aynı çizim aktif olmalıdır.

      Yeşil  → İsim tanımlı alan  (MAHAL BLOCK eşleşti)
      Mavi   → İsimsiz tanımlı alan (polygon var, isim yok)
      Kırmızı→ Tanımsız alan (label var, polygon bulunamadı)

    Args:
        dxf_path: DXF dosyasının tam yolu
    """
    from tools.cad_colorizer import colorize_rooms

    result = colorize_rooms(dxf_path)
    result["summary"] = (
        f"{result['green']} yeşil (isim tanımlı), "
        f"{result['blue']} mavi (isimsiz tanımlı), "
        f"{result['red']} kırmızı (tanımsız)"
    )
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def export_rooms_ifc(
    dxf_path: str,
    output_path: str = "",
    wall_thickness_mm: float = 150.0,
    room_height_m: float = 2.8,
) -> str:
    """
    İŞARETLEME ADIM 4 — Eşleşen mekanları IFC dosyasına aktar.
    match_rooms_to_polygons() verilerini kullanarak her mekan için
    IfcSpace + IfcWall oluşturur. Polygon wall_thickness_mm kadar
    dışarıya genişletilir (iç → dış ölçü dönüşümü).

    Args:
        dxf_path          : Temizlenmiş DXF dosyasının tam yolu
        output_path       : Çıktı IFC yolu (boş bırakılırsa _MEKAN.ifc olarak kaydeder)
        wall_thickness_mm : Duvar kalınlığı mm cinsinden (varsayılan: 150.0 = 15cm)
        room_height_m     : Kat yüksekliği metre cinsinden (varsayılan: 2.8)
    """
    import math as _m, json as _json, uuid as _uuid, time as _time
    import ifcopenshell

    if not output_path:
        base = os.path.splitext(dxf_path)[0]
        output_path = base + "_MEKAN.ifc"

    wall_offset_mm = wall_thickness_mm  # alias for clarity

    # ── 1. Eşleşmiş mekanları al ─────────────────────────────────────────────
    data = _json.loads(match_rooms_to_polygons(dxf_path))
    rooms_raw = [r for r in data["rooms"] if r["status"] != "unmatched" and r["points"]]

    # ── 2. Polygon offset (miter join) ───────────────────────────────────────
    def _offset_polygon(pts_mm, offset_mm):
        n = len(pts_mm)
        if n < 3:
            return pts_mm
        # Sarım yönü: CCW=pozitif alan
        area2 = sum(
            pts_mm[i][0] * pts_mm[(i+1)%n][1] - pts_mm[(i+1)%n][0] * pts_mm[i][1]
            for i in range(n)
        )
        sign = 1 if area2 > 0 else -1   # CCW → dış normal sola, CW → sağa

        result = []
        for i in range(n):
            p0 = pts_mm[(i-1) % n]
            p1 = pts_mm[i]
            p2 = pts_mm[(i+1) % n]

            dx1, dy1 = p1[0]-p0[0], p1[1]-p0[1]
            l1 = _m.hypot(dx1, dy1)
            if l1 < 1e-9:
                result.append(p1)
                continue
            n1x, n1y = -sign*dy1/l1, sign*dx1/l1   # kenar 1'in dış normali

            dx2, dy2 = p2[0]-p1[0], p2[1]-p1[1]
            l2 = _m.hypot(dx2, dy2)
            if l2 < 1e-9:
                result.append(p1)
                continue
            n2x, n2y = -sign*dy2/l2, sign*dx2/l2   # kenar 2'nin dış normali

            bx, by = n1x+n2x, n1y+n2y
            bl = _m.hypot(bx, by)
            if bl < 1e-9:
                bx, by, bl = n1x, n1y, 1.0
            bx, by = bx/bl, by/bl

            cos_h = n1x*bx + n1y*by
            cos_h = max(abs(cos_h), 0.17) * (1 if cos_h >= 0 else -1)  # max 80° miter
            miter = offset_mm / cos_h

            result.append([p1[0] + bx*miter, p1[1] + by*miter])
        return result

    # ── 3. IFC oluştur ────────────────────────────────────────────────────────
    def _uid():
        return ifcopenshell.guid.compress(_uuid.uuid4().hex)

    def _cp3(f, x, y, z=0.):
        return f.createIfcCartesianPoint([float(x), float(y), float(z)])

    def _cp2(f, x, y):
        return f.createIfcCartesianPoint([float(x), float(y)])

    def _dir3(f, x, y, z):
        return f.createIfcDirection([float(x), float(y), float(z)])

    def _ax3(f, origin, zd=None, xd=None):
        return f.createIfcAxis2Placement3D(
            origin,
            zd or _dir3(f, 0, 0, 1),
            xd or _dir3(f, 1, 0, 0),
        )

    def _placement(f, parent, x=0., y=0., z=0.):
        return f.createIfcLocalPlacement(
            parent, _ax3(f, _cp3(f, x, y, z)))

    ifc = ifcopenshell.file(schema="IFC4")
    org   = ifc.createIfcOrganization(None, "CAD Detection", None)
    pers  = ifc.createIfcPerson(None, "Detector", "CAD")
    pao   = ifc.createIfcPersonAndOrganization(pers, org)
    app   = ifc.createIfcApplication(org, "1.0", "CAD Detection Area Definer", "CDAD")
    owner = ifc.createIfcOwnerHistory(pao, app, None, "ADDED", None, pao, app, int(_time.time()))

    ctx  = ifc.createIfcGeometricRepresentationContext(
        None, "Model", 3, 1.0e-5, _ax3(ifc, _cp3(ifc, 0, 0, 0)), None)
    body = ifc.createIfcGeometricRepresentationSubContext(
        "Body", "Model", None, None, None, None, ctx, None, "MODEL_VIEW", None)

    units = ifc.createIfcUnitAssignment([
        ifc.createIfcSIUnit(None, "LENGTHUNIT",     None, "METRE"),
        ifc.createIfcSIUnit(None, "AREAUNIT",       None, "SQUARE_METRE"),
        ifc.createIfcSIUnit(None, "VOLUMEUNIT",     None, "CUBIC_METRE"),
        ifc.createIfcSIUnit(None, "PLANEANGLEUNIT", None, "RADIAN"),
    ])

    project  = ifc.createIfcProject(_uid(), owner, "Mekan IFC", None, None, None, None, None, units)
    site     = ifc.createIfcSite(_uid(), owner, "Site", None, None,
                                 _placement(ifc, None), None, None, "ELEMENT", None)
    building = ifc.createIfcBuilding(_uid(), owner, "Bina", None, None,
                                     _placement(ifc, site.ObjectPlacement),
                                     None, None, "ELEMENT", None, None, None)
    storey   = ifc.createIfcBuildingStorey(_uid(), owner, "Zemin Kat", None, None,
                                     _placement(ifc, building.ObjectPlacement),
                                     None, None, "ELEMENT", 0.0)

    ifc.createIfcRelAggregates(_uid(), owner, None, None, project,  [site])
    ifc.createIfcRelAggregates(_uid(), owner, None, None, site,     [building])
    ifc.createIfcRelAggregates(_uid(), owner, None, None, building, [storey])

    spaces = []
    walls  = []
    MM_TO_M = 0.001

    for idx, room in enumerate(rooms_raw):
        pts_mm  = room["points"]
        pts_off = _offset_polygon(pts_mm, wall_offset_mm)
        pts_m   = [[p[0]*MM_TO_M, p[1]*MM_TO_M] for p in pts_off]

        r_name = room.get("name") or f"Mekan {idx+1}"
        r_num  = room.get("number") or str(idx+1)

        poly_pts = [_cp2(ifc, p[0], p[1]) for p in pts_m]
        poly_pts.append(poly_pts[0])
        profile = ifc.createIfcArbitraryClosedProfileDef(
            "AREA", None, ifc.createIfcPolyline(poly_pts))
        solid = ifc.createIfcExtrudedAreaSolid(
            profile, _ax3(ifc, _cp3(ifc, 0, 0, 0)),
            _dir3(ifc, 0, 0, 1), float(room_height_m))
        shape = ifc.createIfcProductDefinitionShape(None, None, [
            ifc.createIfcShapeRepresentation(body, "Body", "SweptSolid", [solid])])

        space = ifc.createIfcSpace(
            _uid(), owner,
            r_name, r_num, None,
            _placement(ifc, storey.ObjectPlacement),
            shape, None, "ELEMENT", "INTERNAL")

        # Pset_SpaceCommon
        pset_props = [
            ifc.createIfcPropertySingleValue("NetFloorArea", None,
                ifc.createIfcReal(round(room["area_m2"], 3)), None),
            ifc.createIfcPropertySingleValue("MahalAdi", None,
                ifc.createIfcLabel(r_name), None),
            ifc.createIfcPropertySingleValue("MahalNo", None,
                ifc.createIfcLabel(r_num), None),
            ifc.createIfcPropertySingleValue("WallOffsetMM", None,
                ifc.createIfcReal(float(wall_offset_mm)), None),
        ]
        pset = ifc.createIfcPropertySet(_uid(), owner, "Pset_SpaceCommon", None, pset_props)
        ifc.createIfcRelDefinesByProperties(_uid(), owner, None, None, [space], pset)

        spaces.append(space)

        # ── IfcWall: her kenar için ───────────────────────────────────────────
        wall_t_m = wall_thickness_mm * MM_TO_M
        n = len(pts_m)
        for j in range(n):
            p1 = pts_m[j]
            p2 = pts_m[(j + 1) % n]
            dx = p2[0] - p1[0]
            dy = p2[1] - p1[1]
            length = _m.hypot(dx, dy)
            if length < 0.01:
                continue
            angle = _m.atan2(dy, dx)
            w_profile = ifc.createIfcRectangleProfileDef(
                "AREA", None,
                ifc.createIfcAxis2Placement2D(_cp2(ifc, length / 2, wall_t_m / 2), None),
                float(length), float(wall_t_m))
            w_solid = ifc.createIfcExtrudedAreaSolid(
                w_profile, _ax3(ifc, _cp3(ifc, 0, 0, 0)),
                _dir3(ifc, 0, 0, 1), float(room_height_m))
            w_shape = ifc.createIfcProductDefinitionShape(None, None, [
                ifc.createIfcShapeRepresentation(body, "Body", "SweptSolid", [w_solid])])
            wall = ifc.createIfcWall(
                _uid(), owner, f"Wall_R{idx}_S{j}", None, None,
                ifc.createIfcLocalPlacement(
                    storey.ObjectPlacement,
                    _ax3(ifc, _cp3(ifc, p1[0], p1[1], 0.),
                         _dir3(ifc, 0, 0, 1),
                         _dir3(ifc, _m.cos(angle), _m.sin(angle), 0.))),
                w_shape, None, "SOLIDWALL")
            walls.append(wall)

    if spaces:
        ifc.createIfcRelContainedInSpatialStructure(
            _uid(), owner, None, None, spaces, storey)
    if walls:
        ifc.createIfcRelContainedInSpatialStructure(
            _uid(), owner, None, None, walls, storey)

    ifc.write(output_path)

    named   = sum(1 for r in rooms_raw if r.get("name"))
    unnamed = len(rooms_raw) - named

    return json.dumps({
        "mekan_sayisi":       len(spaces),
        "duvar_sayisi":       len(walls),
        "isimli":             named,
        "isimsiz":            unnamed,
        "wall_thickness_mm":  wall_thickness_mm,
        "room_height_m":      room_height_m,
        "output":             output_path,
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def watch_room_at_cursor(dxf_path: str, duration_sec: int = 60) -> str:
    """
    Fare imleci GstarCAD'de hangi mekanın üzerindeyse tooltip gösterir.
    Arka planda çalışır (thread), duration_sec sonra otomatik kapanır.

    Args:
        dxf_path    : Temizlenmiş DXF dosyasının tam yolu
        duration_sec: Kaç saniye çalışacak (varsayılan: 60)
    """
    import threading, time as _time, json as _json, math as _math
    import ctypes, ctypes.wintypes
    import win32api, win32gui
    import win32com.client

    # ── Mekan verilerini önceden yükle ───────────────────────────────────────
    matched   = _json.loads(match_rooms_to_polygons(dxf_path))
    rooms     = [r for r in matched["rooms"] if r["points"]]

    def _in_poly(px, py, pts):
        inside, j = False, len(pts) - 1
        for i in range(len(pts)):
            xi, yi = pts[i]; xj, yj = pts[j]
            if ((yi > py) != (yj > py)) and (px < (xj-xi)*(py-yi)/(yj-yi+1e-12)+xi):
                inside = not inside
            j = i
        return inside

    def _get_draw_coords(acad_doc):
        sx, sy = win32api.GetCursorPos()
        hwnd = win32gui.GetForegroundWindow()
        rect  = win32gui.GetClientRect(hwnd)
        win_w = max(rect[2] - rect[0], 1)
        win_h = max(rect[3] - rect[1], 1)
        pt = ctypes.wintypes.POINT(0, 0)
        ctypes.windll.user32.ClientToScreen(hwnd, ctypes.byref(pt))
        rel_x = (sx - pt.x) / win_w
        rel_y = (sy - pt.y) / win_h
        vp    = acad_doc.ActiveViewport
        vp_cx, vp_cy = vp.Center[0], vp.Center[1]
        vp_h  = vp.Height
        vp_w  = vp_h * (win_w / win_h)
        dx = vp_cx + (rel_x - 0.5) * vp_w
        dy = vp_cy - (rel_y - 0.5) * vp_h
        return dx, dy, sx, sy

    def _run():
        import tkinter as tk
        import pythoncom
        pythoncom.CoInitialize()   # COM thread-safe init

        # GstarCAD bağlantısı
        try:
            acad = win32com.client.GetActiveObject("GstarCAD.Application")
            doc  = acad.ActiveDocument
        except Exception:
            return

        LUMINAIRES   = ["Flat-G", "Lightline 043", "Snow 019"]
        HOTKEYS      = [ord('1'), ord('2'), ord('3')]   # VK codes
        VK_LEFT, VK_RIGHT = 0x25, 0x27                 # ok tuşları
        room_assignments  = {}   # room_id → seçilen armatür
        key_was_down      = [False, False, False]       # 1/2/3 debounce
        arrow_was_down    = [False, False]              # sol/sağ debounce
        current_idx       = [-1]                        # seçili armatür index

        # ── Tooltip penceresi ─────────────────────────────────────────────────
        root = tk.Tk()
        root.overrideredirect(True)
        root.attributes("-topmost", True)
        root.attributes("-alpha", 0.95)
        root.configure(bg="#1a1a2e")
        root.resizable(False, False)

        # Mekan adı satırı
        lbl = tk.Label(
            root, text="", bg="#1a1a2e", fg="#00e5ff",
            font=("Segoe UI", 10, "bold"),
            padx=12, pady=5, anchor="w"
        )
        lbl.pack(fill="x")

        # Ayraç
        tk.Frame(root, bg="#00e5ff", height=1).pack(fill="x", padx=6)

        # Armatür seçim satırı  → [1] Flat-G  [2] Lightline 043  [3] Snow 019
        option_frame = tk.Frame(root, bg="#1a1a2e")
        option_frame.pack(fill="x", padx=6, pady=5)

        opt_labels = []
        for i, name in enumerate(LUMINAIRES):
            frm = tk.Frame(option_frame, bg="#1a1a2e")
            frm.pack(side="left", padx=4)
            num_lbl = tk.Label(frm, text=f"[{i+1}]", bg="#1a1a2e",
                               fg="#ffb300", font=("Segoe UI", 8, "bold"))
            num_lbl.pack(side="left")
            name_lbl = tk.Label(frm, text=f" {name} ", bg="#1a1a2e",
                                fg="#b0bec5", font=("Segoe UI", 9))
            name_lbl.pack(side="left")
            opt_labels.append((frm, num_lbl, name_lbl))

        current_room = [None]

        def _highlight_selection(idx):
            """Seçili armatürü vurgula, diğerlerini sönük göster."""
            for i, (frm, num_lbl, name_lbl) in enumerate(opt_labels):
                if i == idx:
                    frm.config(bg="#0d2137")
                    num_lbl.config(bg="#0d2137", fg="#00e5ff")
                    name_lbl.config(bg="#0d2137", fg="#ffffff")
                else:
                    frm.config(bg="#1a1a2e")
                    num_lbl.config(bg="#1a1a2e", fg="#ffb300")
                    name_lbl.config(bg="#1a1a2e", fg="#b0bec5")

        def _assign_luminaire(idx):
            """Odaya armatür ata."""
            room = current_room[0]
            if room is None:
                return
            idx = idx % len(LUMINAIRES)
            current_idx[0] = idx
            sel      = LUMINAIRES[idx]
            room_id  = room["id"]
            r_name   = room.get("name") or f"Mekan {room_id}"
            r_num    = room.get("number") or ""
            room_assignments[room_id] = {
                "name":      r_name,
                "number":    r_num,
                "luminaire": sel,
            }
            _highlight_selection(idx)
            try:
                prefix = f"{r_num} · " if r_num else ""
                doc.Utility.Prompt(f"\n► {prefix}{r_name} → {sel}\n")
            except Exception:
                pass

        # Başlangıç mesajı
        lbl.config(text="  Mekan bekleniyor...  ")
        sx0, sy0 = win32api.GetCursorPos()
        root.geometry(f"+{sx0 + 16}+{sy0 + 16}")
        root.update()

        last_room_id   = [None]
        hover_ent      = [None]   # GstarCAD'deki kırmızı çerçeve entity objesi
        end_time       = _time.time() + duration_sec
        HOVER_LAYER    = "AI_HOVER"

        # AI_HOVER layer oluştur (kırmızı)
        try:
            hl = doc.Layers.Add(HOVER_LAYER)
            hl.Color = 1   # kırmızı
        except Exception:
            pass

        def _draw_highlight(pts_mm):
            """Odanın polygon'unu GstarCAD'de kırmızı çizer, entity objesi döner."""
            try:
                import pythoncom as _pc
                flat = []
                for p in pts_mm:
                    flat += [p[0], p[1]]
                pts_var = win32com.client.VARIANT(
                    _pc.VT_ARRAY | _pc.VT_R8, flat)
                msp = doc.ModelSpace
                poly = msp.AddLightWeightPolyline(pts_var)
                poly.Closed = True
                poly.Layer  = HOVER_LAYER
                poly.Color  = 1       # kırmızı
                poly.LineWeight = 50  # 0.50mm kalın
                return poly
            except Exception:
                return None

        def _erase_highlight(ent):
            """Entity objesini sil."""
            if ent is None:
                return
            try:
                ent.Delete()
            except Exception:
                pass

        def _tick():
            if _time.time() > end_time:
                _erase_highlight(hover_ent[0])
                root.destroy()
                return

            try:
                # GstarCAD penceresini bul
                gcad_hwnd = [None]
                def _find(h, _):
                    if not gcad_hwnd[0]:
                        try:
                            t = win32gui.GetWindowText(h)
                            c = win32gui.GetClassName(h)
                            if win32gui.IsWindowVisible(h) and ("GstarCAD" in t or "GCAD" in c):
                                gcad_hwnd[0] = h
                        except Exception:
                            pass
                win32gui.EnumWindows(_find, None)
                hwnd = gcad_hwnd[0] or win32gui.GetForegroundWindow()

                rect  = win32gui.GetClientRect(hwnd)
                win_w = max(rect[2] - rect[0], 1)
                win_h = max(rect[3] - rect[1], 1)
                pt    = ctypes.wintypes.POINT(0, 0)
                ctypes.windll.user32.ClientToScreen(hwnd, ctypes.byref(pt))

                sx, sy = win32api.GetCursorPos()
                rel_x  = (sx - pt.x) / win_w
                rel_y  = (sy - pt.y) / win_h

                vp    = doc.ActiveViewport
                vp_cx, vp_cy = vp.Center[0], vp.Center[1]
                vp_h  = vp.Height
                vp_w  = vp_h * (win_w / win_h)
                dx    = vp_cx + (rel_x - 0.5) * vp_w
                dy    = vp_cy - (rel_y - 0.5) * vp_h

                found = None
                for room in rooms:
                    if _in_poly(dx, dy, room["points"]):
                        found = room
                        break

                # ── Hotkey poll: 1/2/3 ve ←/→ ok tuşları ───────────────────
                for ki, vk in enumerate(HOTKEYS):
                    is_down = bool(win32api.GetAsyncKeyState(vk) & 0x8000)
                    if is_down and not key_was_down[ki]:
                        _assign_luminaire(ki)
                    key_was_down[ki] = is_down

                # Sol ok → önceki armatür
                left_down = bool(win32api.GetAsyncKeyState(VK_LEFT) & 0x8000)
                if left_down and not arrow_was_down[0] and current_room[0] is not None:
                    _assign_luminaire(current_idx[0] - 1)
                arrow_was_down[0] = left_down

                # Sağ ok → sonraki armatür
                right_down = bool(win32api.GetAsyncKeyState(VK_RIGHT) & 0x8000)
                if right_down and not arrow_was_down[1] and current_room[0] is not None:
                    _assign_luminaire(current_idx[0] + 1)
                arrow_was_down[1] = right_down

                if found:
                    rid = found["id"]
                    if rid != last_room_id[0]:
                        # Eski çerçeveyi sil, yenisini çiz
                        _erase_highlight(hover_ent[0])
                        hover_ent[0]    = _draw_highlight(found["points"])
                        last_room_id[0] = rid
                        current_room[0] = found
                        name   = found.get("name") or "İSİMSİZ"
                        number = found.get("number") or ""
                        area   = found.get("area_m2", 0)
                        prefix = f"{number} · " if number else ""
                        lbl.config(text=f"  {prefix}{name}  |  {area:.1f} m²")
                        # Önceki seçimi geri yükle (varsa)
                        prev_lum = room_assignments.get(rid, {}).get("luminaire", "")
                        if prev_lum in LUMINAIRES:
                            prev_idx = LUMINAIRES.index(prev_lum)
                            current_idx[0] = prev_idx
                            _highlight_selection(prev_idx)
                        else:
                            current_idx[0] = -1
                            _highlight_selection(-1)
                    root.geometry(f"+{sx + 16}+{sy + 16}")
                    root.deiconify()
                else:
                    if last_room_id[0] is not None:
                        _erase_highlight(hover_ent[0])
                        hover_ent[0]    = None
                        last_room_id[0] = None
                        current_room[0] = None
                    root.withdraw()

            except Exception:
                pass

            root.after(200, _tick)

        root.after(200, _tick)
        root.mainloop()
        _erase_highlight(hover_ent[0])
        # Atamalar sonucu JSON'a yaz
        import json as _j
        out_path = os.path.splitext(dxf_path)[0] + "_armatür.json"
        try:
            with open(out_path, "w", encoding="utf-8") as _f:
                _j.dump(room_assignments, _f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        pythoncom.CoUninitialize()

    t = threading.Thread(target=_run, daemon=True)
    t.start()

    return _json.dumps({
        "durum":      "başlatıldı",
        "sure_sn":    duration_sec,
        "mekan_sayisi": len(rooms),
        "mesaj":      f"Fare tooltip {duration_sec} saniye aktif. GstarCAD'de odalara göz atın.",
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def get_room_at_cursor(dxf_path: str) -> str:
    """
    Fare imlecinin GstarCAD'de üzerinde bulunduğu mekanı döner.
    GstarCAD aktif ve dxf_path ile aynı çizim açık olmalıdır.

    1. Win32 API ile ekran koordinatını alır
    2. GstarCAD COM viewport'undan çizim koordinatına çevirir
    3. match_rooms_to_polygons() polygon listesine karşı point-in-polygon uygular
    4. Mekan adı + numarası + alanı döner

    Args:
        dxf_path: Temizlenmiş DXF dosyasının tam yolu
    """
    import win32api, win32gui, win32con
    import win32com.client, pythoncom
    import json as _json, math as _math, time as _time
    import ctypes

    # ── 1. GstarCAD bağlantısı ───────────────────────────────────────────────
    for attempt in range(3):
        try:
            acad = win32com.client.GetActiveObject("GstarCAD.Application")
            doc  = acad.ActiveDocument
            _ = doc.Name
            break
        except Exception:
            if attempt == 2:
                return _json.dumps({"hata": "GstarCAD bağlantısı kurulamadı"}, ensure_ascii=False)
            _time.sleep(1)

    # ── 2. Ekran fare koordinatı ─────────────────────────────────────────────
    screen_x, screen_y = win32api.GetCursorPos()

    # ── 3. GstarCAD pencere bilgileri ────────────────────────────────────────
    # GstarCAD ana penceresi
    hwnd = None
    def _enum(h, _):
        nonlocal hwnd
        cls = win32gui.GetClassName(h)
        if "GstarCAD" in cls or "GCAD" in cls or "AutoCAD" in cls:
            hwnd = h
    win32gui.EnumWindows(_enum, None)

    if hwnd is None:
        # Fallback: foreground window
        hwnd = win32gui.GetForegroundWindow()

    rect = win32gui.GetClientRect(hwnd)
    win_w = rect[2] - rect[0]
    win_h = rect[3] - rect[1]

    # Pencere sol-üst köşesi (screen koordinatında)
    client_origin = ctypes.wintypes.POINT(0, 0)
    ctypes.windll.user32.ClientToScreen(hwnd, ctypes.byref(client_origin))
    win_left = client_origin.x
    win_top  = client_origin.y

    # Fare → pencere içi pozisyon (0..1 arası normalize)
    rel_x = (screen_x - win_left) / max(win_w, 1)
    rel_y = (screen_y - win_top)  / max(win_h, 1)

    # ── 4. Viewport'tan çizim koordinatına çevir ─────────────────────────────
    try:
        vp      = doc.ActiveViewport
        vp_cx   = vp.Center[0]   # çizim merkezi x
        vp_cy   = vp.Center[1]   # çizim merkezi y
        vp_h    = vp.Height       # viewport yüksekliği (çizim birimi)
        vp_w    = vp_h * (win_w / max(win_h, 1))   # en-boy oranı koru

        draw_x = vp_cx + (rel_x - 0.5) * vp_w
        draw_y = vp_cy - (rel_y - 0.5) * vp_h   # Y ekranı ters
    except Exception as e:
        return _json.dumps({"hata": f"Viewport bilgisi alınamadı: {e}"}, ensure_ascii=False)

    # ── 5. Point-in-polygon ───────────────────────────────────────────────────
    matched = _json.loads(match_rooms_to_polygons(dxf_path))
    rooms   = [r for r in matched["rooms"] if r["points"]]

    def _in_poly(px, py, pts):
        inside, j = False, len(pts) - 1
        for i in range(len(pts)):
            xi, yi = pts[i]
            xj, yj = pts[j]
            if ((yi > py) != (yj > py)) and (px < (xj-xi)*(py-yi)/(yj-yi+1e-12)+xi):
                inside = not inside
            j = i
        return inside

    found = None
    for room in rooms:
        if _in_poly(draw_x, draw_y, room["points"]):
            found = room
            break

    if found:
        return _json.dumps({
            "bulundu":   True,
            "name":      found["name"],
            "number":    found["number"],
            "area_m2":   found["area_m2"],
            "status":    found["status"],
            "cursor_draw": [round(draw_x, 1), round(draw_y, 1)],
        }, ensure_ascii=False, indent=2)
    else:
        return _json.dumps({
            "bulundu":     False,
            "mesaj":       "İmleç herhangi bir mekanın üzerinde değil",
            "cursor_draw": [round(draw_x, 1), round(draw_y, 1)],
        }, ensure_ascii=False, indent=2)


@mcp.tool()
def open_luminaire_picker(
    dxf_path: str,
    luminaires: str = "Flat-G,Lightline 043,Snow 019",
) -> str:
    """
    Mekan → Armatür atama penceresi açar.
    Sol panel: tüm mekanlar listesi.
    Sağ panel: armatür listesi.
    Mekan seçilince GstarCAD'de kırmızı çerçeve + zoom.
    Armatür tıklanınca atanır. Kaydet → JSON dosyası.

    Args:
        dxf_path  : Temizlenmiş DXF dosyasının tam yolu
        luminaires: Virgülle ayrılmış armatür listesi
    """
    import threading, json as _json, time as _time, math as _m
    import win32com.client, pythoncom

    lum_list = [l.strip() for l in luminaires.split(",") if l.strip()]

    # Mekan verilerini yükle
    data  = _json.loads(match_rooms_to_polygons(dxf_path))
    rooms = [r for r in data["rooms"] if r["points"]]  # polygon'u olanlar

    # Çizim birimini oku → 600mm armatür boyutunu drawing unit'e çevir
    try:
        import ezdxf as _ezdxf
        _dxf_doc = _ezdxf.readfile(dxf_path, encoding='utf-8')
        _insunits = _dxf_doc.header.get('$INSUNITS', 0)
    except Exception:
        _insunits = 0
    # INSUNITS: 4=mm, 5=cm, 6=m, 1=inch, 0=birim yok (mm varsay)
    _MM_PER_UNIT = {4: 1.0, 5: 10.0, 6: 1000.0, 1: 25.4, 2: 304.8, 3: 1609344.0}
    _mm_per = _MM_PER_UNIT.get(_insunits, 1.0)   # 1 drawing unit = kaç mm
    BLOCK_SIZE  = 600.0 / _mm_per    # 600mm armatür boyutu → drawing units

    # Armatür lümen değerleri (default) — isme göre eşleştirme
    _LUM_LUMENS_DEFAULT = {
        "flat-g":       3600,
        "lightline 043":4300,
        "lightline":    4300,
        "snow 019":     2800,
        "snow":         2800,
    }
    def _lum_lumens(name: str) -> int:
        k = name.lower()
        for pat, val in _LUM_LUMENS_DEFAULT.items():
            if pat in k:
                return val
        return 3000  # bilinmeyen armatür için default

    # Oda tipine göre hedef lüx tablosu (Türkçe isim eşleştirme)
    _ROOM_LUX = [
        (["laboratuvar", "lab"],          750),
        (["muayene", "tedavi", "ameliyat"], 500),
        (["ofis", "büro", "çalışma"],      500),
        (["mutfak", "yemekhane"],          300),
        (["wc", "tuvalet", "banyo", "duş"],200),
        (["koridor", "hol", "giriş", "lobi"], 200),
        (["depo", "arşiv", "teknik"],      150),
        (["mescit", "ibadet"],             300),
    ]
    def _room_lux(room_name: str) -> int:
        k = (room_name or "").lower()
        for keywords, lux in _ROOM_LUX:
            if any(kw in k for kw in keywords):
                return lux
        return 300  # genel ofis/mekan default

    def _run():
        import tkinter as tk
        import tkinter.ttk as ttk
        import tkinter.messagebox as msgbox
        pythoncom.CoInitialize()

        # GstarCAD bağlantısı
        try:
            acad = win32com.client.GetActiveObject("GstarCAD.Application")
            doc  = acad.ActiveDocument
            msp  = doc.ModelSpace
        except Exception:
            acad = doc = msp = None

        HOVER_LAYER = "AI_HOVER"
        if doc:
            try:
                hl = doc.Layers.Add(HOVER_LAYER)
                hl.Color = 1
            except Exception:
                pass

        assignments = {}   # room_id → luminaire
        hover_ent   = [None]

        def _draw_border(pts_mm):
            if not doc:
                return None
            try:
                flat = []
                for p in pts_mm:
                    flat += [p[0], p[1]]
                pts_var = win32com.client.VARIANT(
                    pythoncom.VT_ARRAY | pythoncom.VT_R8, flat)
                poly = msp.AddLightWeightPolyline(pts_var)
                poly.Closed   = True
                poly.Layer    = HOVER_LAYER
                poly.Color    = 1
                poly.LineWeight = 50
                return poly
            except Exception:
                return None

        def _erase_border(ent):
            if ent:
                try:
                    ent.Delete()
                except Exception:
                    pass

        def _zoom_to(room):
            if not doc:
                return
            try:
                pts = room["points"]
                xs  = [p[0] for p in pts]
                ys  = [p[1] for p in pts]
                pad = 2000
                doc.SendCommand(
                    f"ZOOM W {min(xs)-pad},{min(ys)-pad} {max(xs)+pad},{max(ys)+pad}\n")
            except Exception:
                pass

        # Oda → çizilen armatür entity listesi  {room_id: [ent, ent, ...]}
        placed_ents = {}

        def _in_poly_local(px, py, pts):
            inside, j = False, len(pts) - 1
            for i in range(len(pts)):
                xi, yi = pts[i]; xj, yj = pts[j]
                if ((yi > py) != (yj > py)) and (px < (xj-xi)*(py-yi)/(yj-yi+1e-12)+xi):
                    inside = not inside
                j = i
            return inside

        def _erase_placed(room_id):
            for ent in placed_ents.get(room_id, []):
                try:
                    ent.Delete()
                except Exception:
                    pass
            placed_ents[room_id] = []

        def _place_luminaires(room, lum_name):
            """Lüx hesabına göre armatür grid yerleşimi — birim otomatik."""
            import math as _math
            if not doc:
                return 0, "GstarCAD bağlantısı yok"
            room_id = room["id"]
            _erase_placed(room_id)

            pts = room["points"]
            xs  = [p[0] for p in pts]
            ys  = [p[1] for p in pts]
            min_x, max_x = min(xs), max(xs)
            min_y, max_y = min(ys), max(ys)
            w = max_x - min_x   # drawing units
            h = max_y - min_y

            # Alan m² — drawing units → mm → m²
            area_mm2 = (w * _mm_per) * (h * _mm_per)
            area_m2  = area_mm2 / 1_000_000

            # Lüx hesabı: N = ceil(E × A / (Φ × MLF × UF))
            E   = lux_var.get()
            MLF = mlf_var.get()
            UF  = uf_var.get()
            phi = _lum_lumens(lum_name)
            N   = max(1, _math.ceil((E * area_m2) / (phi * MLF * UF)))

            # Grid düzeni: oda en/boy oranına uygun
            aspect = (w / h) if h > 0 else 1.0
            n_y = max(1, round(_math.sqrt(N / aspect)))
            n_x = max(1, _math.ceil(N / n_y))

            sp_x = w / n_x
            sp_y = h / n_y
            half = BLOCK_SIZE / 2

            # Bilgi etiketi güncelle
            lux_info.config(text=f"N={N}  ({n_x}×{n_y} grid)  Φ={phi} lm  A={area_m2:.1f}m²")

            ARM_LAYER = "ARMATÜR"
            try:
                cur_msp = doc.ActiveDocument.ModelSpace if hasattr(doc, 'ActiveDocument') else msp
            except Exception:
                cur_msp = msp

            try:
                al = cur_msp.Document.Layers.Add(ARM_LAYER)
                al.Color = 3
            except Exception:
                try:
                    doc.Layers.Add(ARM_LAYER).Color = 3
                except Exception:
                    pass

            placed = []
            errors = []
            for i in range(n_x):
                for j in range(n_y):
                    cx = min_x + sp_x * (i + 0.5)
                    cy = min_y + sp_y * (j + 0.5)
                    if not _in_poly_local(cx, cy, pts):
                        continue
                    try:
                        flat = [cx-half, cy-half,
                                cx+half, cy-half,
                                cx+half, cy+half,
                                cx-half, cy+half]
                        pts_var = win32com.client.VARIANT(
                            pythoncom.VT_ARRAY | pythoncom.VT_R8, flat)
                        sq = msp.AddLightWeightPolyline(pts_var)
                        sq.Closed     = True
                        sq.Layer      = ARM_LAYER
                        sq.Color      = 3
                        sq.LineWeight = 25
                        placed.append(sq)
                    except Exception as e:
                        errors.append(str(e))

            # Armatür adı etiketi
            try:
                cx_c  = (min_x + max_x) / 2
                cy_c  = (min_y + max_y) / 2
                tp    = win32com.client.VARIANT(
                    pythoncom.VT_ARRAY | pythoncom.VT_R8, [cx_c, cy_c, 0.0])
                txt_h = max(min(w * 0.03, 100), 25)
                t = msp.AddText(lum_name, tp, txt_h)
                t.Layer = ARM_LAYER
                t.Color = 3
                placed.append(t)
            except Exception as e:
                errors.append(f"text:{e}")

            placed_ents[room_id] = placed
            err_str = (" | HATA: " + errors[0]) if errors else ""
            return len(placed), err_str

        # ── Tkinter penceresi ─────────────────────────────────────────────
        root = tk.Tk()
        _unit_label = {4:"mm", 5:"cm", 6:"m", 1:"inç", 0:"mm(?)"}.get(_insunits, str(_insunits))
        root.title(f"Mekan Armatür Atama  [birim: {_unit_label} · armatür: {BLOCK_SIZE:.0f}×{BLOCK_SIZE:.0f} {_unit_label}]")
        root.configure(bg="#0d1b2a")
        root.geometry("680x520")
        root.resizable(True, True)
        root.attributes("-topmost", True)

        BG      = "#0d1b2a"
        BG2     = "#1a2a3a"
        ACCENT  = "#00e5ff"
        FG      = "#e0e0ff"
        FG2     = "#b0bec5"
        SEL_BG  = "#0d3a4a"
        LUM_SEL = "#1a4a1a"
        FONT    = ("Segoe UI", 10)
        FONT_B  = ("Segoe UI", 10, "bold")
        FONT_S  = ("Segoe UI", 9)

        # ── Başlık ───────────────────────────────────────────────────────
        hdr = tk.Label(root, text="MEKAN ARMATÜR ATAMA",
                       bg=BG, fg=ACCENT, font=("Segoe UI", 12, "bold"), pady=8)
        hdr.pack(fill="x")
        tk.Frame(root, bg=ACCENT, height=1).pack(fill="x")

        # ── Lüx kontrol paneli ────────────────────────────────────────────
        lux_frame = tk.Frame(root, bg="#0a1520", pady=4)
        lux_frame.pack(fill="x", padx=10)
        tk.Label(lux_frame, text="Hedef Lüx:", bg="#0a1520", fg=FG2,
                 font=FONT_S).pack(side="left")
        lux_var = tk.IntVar(value=300)
        lux_spin = tk.Spinbox(lux_frame, from_=50, to=2000, increment=50,
                              textvariable=lux_var, width=6,
                              bg="#1a2a3a", fg=ACCENT, font=FONT_S,
                              relief="flat", bd=1)
        lux_spin.pack(side="left", padx=(4,12))
        tk.Label(lux_frame, text="lx  |  MLF:", bg="#0a1520", fg=FG2,
                 font=FONT_S).pack(side="left")
        mlf_var = tk.DoubleVar(value=0.80)
        tk.Spinbox(lux_frame, from_=0.5, to=1.0, increment=0.05,
                   textvariable=mlf_var, format="%.2f", width=5,
                   bg="#1a2a3a", fg=ACCENT, font=FONT_S,
                   relief="flat", bd=1).pack(side="left", padx=(4,12))
        tk.Label(lux_frame, text="UF:", bg="#0a1520", fg=FG2,
                 font=FONT_S).pack(side="left")
        uf_var = tk.DoubleVar(value=0.65)
        tk.Spinbox(lux_frame, from_=0.3, to=1.0, increment=0.05,
                   textvariable=uf_var, format="%.2f", width=5,
                   bg="#1a2a3a", fg=ACCENT, font=FONT_S,
                   relief="flat", bd=1).pack(side="left", padx=(4,12))
        lux_info = tk.Label(lux_frame, text="", bg="#0a1520", fg="#ffcc44", font=FONT_S)
        lux_info.pack(side="left", padx=8)

        # ── Ana içerik ───────────────────────────────────────────────────
        content = tk.Frame(root, bg=BG)
        content.pack(fill="both", expand=True, padx=10, pady=8)

        # Sol: mekan listesi
        left = tk.Frame(content, bg=BG)
        left.pack(side="left", fill="both", expand=True)
        tk.Label(left, text="MEKANLAR", bg=BG, fg=ACCENT,
                 font=FONT_B).pack(anchor="w", pady=(0,4))

        room_lb = tk.Listbox(left, bg=BG2, fg=FG, selectbackground=SEL_BG,
                             selectforeground=ACCENT, font=FONT_S,
                             relief="flat", bd=0, activestyle="none",
                             highlightthickness=1, highlightcolor=ACCENT,
                             highlightbackground=BG2)
        sb_r = ttk.Scrollbar(left, orient="vertical", command=room_lb.yview)
        room_lb.config(yscrollcommand=sb_r.set)
        sb_r.pack(side="right", fill="y")
        room_lb.pack(fill="both", expand=True)

        # Mekan listesini doldur
        for r in rooms:
            num    = r.get("number") or ""
            name   = r.get("name") or "İSİMSİZ"
            area   = r.get("area_m2", 0)
            prefix = f"{num} · " if num else ""
            room_lb.insert("end", f"  {prefix}{name}  [{area:.1f} m²]")

        # Orta ayraç
        tk.Frame(content, bg=ACCENT, width=1).pack(side="left", fill="y", padx=8)

        # Sağ: armatür listesi
        right = tk.Frame(content, bg=BG)
        right.pack(side="left", fill="both", expand=True)
        tk.Label(right, text="ARMATÜRLER", bg=BG, fg=ACCENT,
                 font=FONT_B).pack(anchor="w", pady=(0,4))

        lum_lb = tk.Listbox(right, bg=BG2, fg=FG, selectbackground=LUM_SEL,
                            selectforeground="#00ff88", font=FONT,
                            relief="flat", bd=0, activestyle="none",
                            highlightthickness=1, highlightcolor=ACCENT,
                            highlightbackground=BG2)
        for lum in lum_list:
            lum_lb.insert("end", f"  {lum}")
        lum_lb.pack(fill="both", expand=True)

        # ── Durum satırı ─────────────────────────────────────────────────
        tk.Frame(root, bg=ACCENT, height=1).pack(fill="x")
        status_lbl = tk.Label(root, text="Mekan seçin, ardından armatür tıklayın.",
                              bg=BG, fg=FG2, font=FONT_S, pady=4)

        def _set_status(msg):
            status_lbl.config(text=msg)
        status_lbl.pack(fill="x", padx=10)

        # ── Alt butonlar ─────────────────────────────────────────────────
        btn_frame = tk.Frame(root, bg=BG)
        btn_frame.pack(fill="x", padx=10, pady=6)

        assigned_lbl = tk.Label(btn_frame, text="Atanan: 0", bg=BG, fg=FG2,
                                font=FONT_S)
        assigned_lbl.pack(side="left")

        def _save():
            out = os.path.splitext(dxf_path)[0] + "_armatür.json"
            with open(out, "w", encoding="utf-8") as f:
                _json.dump(assignments, f, ensure_ascii=False, indent=2)
            msgbox.showinfo("Kaydedildi", f"{len(assignments)} atama kaydedildi.\n{out}")

        selected_room = [None]

        def _on_room_select(event=None):
            sel = room_lb.curselection()
            if not sel:
                return
            idx  = sel[0]
            room = rooms[idx]
            selected_room[0] = room

            # Eski border sil, yeni çiz + zoom
            _erase_border(hover_ent[0])
            hover_ent[0] = _draw_border(room["points"])
            _zoom_to(room)

            # Mevcut atamayı göster
            rid      = room["id"]
            cur_lum  = assignments.get(rid, {}).get("luminaire", "")
            if cur_lum in lum_list:
                lum_lb.selection_clear(0, "end")
                lum_lb.selection_set(lum_list.index(cur_lum))
            name = room.get("name") or "İSİMSİZ"
            num  = room.get("number") or ""
            prefix = f"{num} · " if num else ""
            # Oda tipine göre önerilen lüx değerini lux_var'a yaz
            suggested = _room_lux(name)
            lux_var.set(suggested)
            _set_status(f"Seçili: {prefix}{name}  — Önerilen lüx: {suggested}  — Armatür seçin")

        selected_lum = [None]   # seçili armatür — butona tıkta kaybolmasın

        def _on_lum_select(event=None):
            sel_l = lum_lb.curselection()
            if not sel_l:
                return
            selected_lum[0] = lum_list[sel_l[0]]
            room = selected_room[0]
            if room:
                r_name = room.get("name") or f"Mekan {room['id']}"
                r_num  = room.get("number") or ""
                prefix = f"{r_num} · " if r_num else ""
                _set_status(f"Seçili: {prefix}{r_name}  →  {selected_lum[0]}  — Çiz'e bas")
            else:
                _set_status(f"Armatür: {selected_lum[0]}  — Mekan seçin")

        def _on_draw():
            room = selected_room[0]
            lum  = selected_lum[0]
            if not room or not lum:
                _set_status("Önce mekan ve armatür seçin!")
                return
            rid    = room["id"]
            r_name = room.get("name") or f"Mekan {rid}"
            r_num  = room.get("number") or ""
            assignments[rid] = {"name": r_name, "number": r_num, "luminaire": lum}
            assigned_lbl.config(text=f"Atanan: {len(assignments)}")
            prefix = f"{r_num} · " if r_num else ""
            count, err = _place_luminaires(room, lum)
            _set_status(f"✓  {prefix}{r_name}  →  {lum}  ({count} adet){err}")
            # Listede yeşil vurgula
            try:
                idx = rooms.index(room)
                room_lb.itemconfig(idx, fg="#00ff88")
            except ValueError:
                pass

        def _on_close():
            _erase_border(hover_ent[0])
            root.destroy()

        room_lb.bind("<<ListboxSelect>>", _on_room_select)
        lum_lb.bind("<<ListboxSelect>>", _on_lum_select)
        lum_lb.bind("<Double-Button-1>", lambda e: _on_draw())

        # ── Butonlar (fonksiyonlar tanımlandıktan sonra) ──────────────────
        tk.Button(btn_frame, text="Kaydet", bg="#1a3a1a", fg="#00ff88",
                  font=FONT_B, relief="flat", padx=16, pady=4,
                  command=_save).pack(side="right")
        tk.Button(btn_frame, text="Çiz", bg="#1a2a3a", fg="#00cfff",
                  font=FONT_B, relief="flat", padx=16, pady=4,
                  command=_on_draw).pack(side="right", padx=6)
        tk.Button(btn_frame, text="Kapat", bg="#3a1a1a", fg="#ff6b6b",
                  font=FONT_S, relief="flat", padx=12, pady=4,
                  command=_on_close).pack(side="right", padx=6)

        root.protocol("WM_DELETE_WINDOW", _on_close)
        root.mainloop()
        pythoncom.CoUninitialize()

    _err_box = [None]

    def _run_safe():
        try:
            _run()
        except Exception as _e:
            import traceback
            _err_box[0] = traceback.format_exc()

    t = threading.Thread(target=_run_safe, daemon=True)
    t.start()
    t.join(timeout=3.0)   # 3 sn bekle — hata varsa yakala

    if _err_box[0]:
        return _json.dumps({"hata": _err_box[0]}, ensure_ascii=False)

    return _json.dumps({
        "durum":   "pencere açıldı",
        "mekanlar": len(rooms),
        "armatürler": lum_list,
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def clean_lighting(dxf_path: str, output_path: str = "") -> str:
    """
    ADIM 1 — Aydınlatma temizleme.
    DXF'ten tüm aydınlatma armatürlerini ve hatch taramalarını kaldırır.
    Duvarları açığa çıkarmak için ilk adımdır.

    Kaldırılanlar:
      • Aydınlatma layer'larındaki tüm INSERT (armatür sembolleri)
      • Modelspace'deki tüm HATCH (duvar/kolon taramaları)
      • Blok tanımları içindeki HATCH (mimari bloklar)

    Args:
        dxf_path   : Kaynak DXF dosyasının tam yolu
        output_path: Çıktı yolu (boş bırakılırsa _TEMIZ suffix eklenir)

    Döner:
        Silinen entity sayıları ve çıktı dosya yolu
    """
    import ezdxf as _ezdxf
    from pathlib import Path

    _LIGHTING_LAYERS = {
        "E-SEMBOL", "ELKSEMBOL", "ELEKTRIK", "KZY.SEMBOL",
        "MYSEMBOL", "KZY-SEMBOL", "KZY.AYDINLATMA",
        "MYAYDINLATMA", "B_AYDINLATMA", "1-ARMATUR"
    }

    def _is_lighting_layer(lu: str) -> bool:
        if "AYDINLATMA" in lu or "ARMATUR" in lu or "ARMATÜR" in lu:
            return True
        if "HAT" in lu or "KESIT" in lu:
            return False
        return lu in _LIGHTING_LAYERS

    doc = _ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    # 1. Armatür INSERT'lerini sil
    armatur_del = []
    for e in msp:
        if _is_protected(e.dxf.layer):
            continue
        if e.dxftype() == "INSERT" and _is_lighting_layer(e.dxf.layer.upper()):
            armatur_del.append(e)
    for e in armatur_del:
        msp.delete_entity(e)

    # 2. Modelspace HATCH sil
    hatch_msp = [e for e in msp if e.dxftype() == "HATCH"]
    for e in hatch_msp:
        msp.delete_entity(e)

    # 3. Blok tanımları içindeki HATCH sil (armatür blokları hariç)
    armatur_block_names = {e.dxf.name for e in msp if e.dxftype() == "INSERT"}
    hatch_block = 0
    for block in doc.blocks:
        if block.name in armatur_block_names:
            continue
        hatches = [e for e in block if e.dxftype() == "HATCH"]
        for e in hatches:
            block.delete_entity(e)
        hatch_block += len(hatches)

    if not output_path:
        p = Path(dxf_path)
        output_path = str(p.parent / (p.stem + "_TEMIZ" + p.suffix))

    doc.saveas(output_path)

    result = {
        "armatür_silindi": len(armatur_del),
        "hatch_modelspace_silindi": len(hatch_msp),
        "hatch_blok_silindi": hatch_block,
        "toplam_silindi": len(armatur_del) + len(hatch_msp) + hatch_block,
        "cikti_dosya": output_path,
        "sonraki_adim": "clean_cables() ile kablo hatlarını temizleyin"
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def clean_cables(dxf_path: str, output_path: str = "") -> str:
    """
    ADIM 2 — Kablo/hat temizleme.
    clean_lighting() sonrasında çalıştırılır.
    Kablo, priz ve toplama hattı layer'larındaki entity'leri kaldırır.

    Kaldırılanlar:
      • *HAT* layer'larındaki LINE ve LWPOLYLINE (aydınlatma/priz/toplama hatları)
      • *PRIZ* layer'larındaki INSERT (priz sembolleri)

    Args:
        dxf_path   : Kaynak DXF (genellikle clean_lighting çıktısı)
        output_path: Çıktı yolu (boş bırakılırsa _KABLO suffix eklenir)
    """
    import ezdxf as _ezdxf
    from pathlib import Path

    def _is_cable_layer(lu: str) -> bool:
        return ("HAT" in lu or "KABLO" in lu or "CABLE" in lu
                or "PRIZ" in lu or "TOPRAK" in lu or "TOPLAMA" in lu)

    doc = _ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    kablo_del = []
    for e in msp:
        layer = e.dxf.layer
        if _is_protected(layer):
            continue
        lu = layer.upper()
        if _is_cable_layer(lu) and e.dxftype() in ("LINE", "LWPOLYLINE", "INSERT", "MTEXT", "TEXT"):
            kablo_del.append(e)
    for e in kablo_del:
        msp.delete_entity(e)

    if not output_path:
        p = Path(dxf_path)
        output_path = str(p.parent / (p.stem + "_KABLO" + p.suffix))

    doc.saveas(output_path)

    result = {
        "kablo_silindi": len(kablo_del),
        "cikti_dosya": output_path,
        "sonraki_adim": "detect_rooms() ile alan tespiti yapabilirsiniz"
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def clean_block_hatches(dxf_path: str, output_path: str = "") -> str:
    """
    ADIM 3 — Blok referans içindeki hatch temizleme.
    clean_cables() sonrasında çalıştırılır.
    Modelspace'deki INSERT'lerin referans ettiği blok tanımları
    içindeki tüm HATCH entity'lerini kaldırır.

    Hedef bloklar: asma tavan, logo, mimari detay bloklarındaki taramalar.
    Armatür blokları (aydınlatma layer'ında olanlar) korunur.

    Args:
        dxf_path   : Kaynak DXF (genellikle clean_cables çıktısı)
        output_path: Çıktı yolu (boş bırakılırsa _BLOK suffix eklenir)
    """
    import ezdxf as _ezdxf
    from pathlib import Path

    def _is_lighting_layer(lu: str) -> bool:
        if "AYDINLATMA" in lu or "ARMATUR" in lu or "ARMATÜR" in lu:
            return True
        if "HAT" in lu or "KESIT" in lu:
            return False
        return lu in {"E-SEMBOL", "ELKSEMBOL", "ELEKTRIK", "KZY.SEMBOL",
                      "MYSEMBOL", "KZY-SEMBOL", "1-ARMATUR"}

    doc = _ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    cleaned_blocks = {}
    total = 0

    for e in msp:
        if e.dxftype() != "INSERT":
            continue
        # Armatür layer'ındaki INSERT'lerin bloklarını koru
        if _is_lighting_layer(e.dxf.layer.upper()):
            continue
        block_name = e.dxf.name
        if block_name in cleaned_blocks:
            continue
        try:
            block = doc.blocks.get(block_name)
            if block is None:
                continue
            hatches = [x for x in block if x.dxftype() == "HATCH"]
            for h in hatches:
                block.delete_entity(h)
            if hatches:
                cleaned_blocks[block_name] = len(hatches)
                total += len(hatches)
        except Exception:
            pass

    if not output_path:
        p = Path(dxf_path)
        output_path = str(p.parent / (p.stem + "_BLOK" + p.suffix))

    doc.saveas(output_path)

    result = {
        "temizlenen_blok_sayisi": len(cleaned_blocks),
        "hatch_silindi": total,
        "bloklar": cleaned_blocks,
        "cikti_dosya": output_path,
        "sonraki_adim": "detect_rooms() ile alan tespiti yapabilirsiniz"
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def clean_hatch(dxf_path: str, output_path: str = "") -> str:
    """
    ADIM 4 — Modelspace'deki tüm HATCH temizleme.
    Pipeline'dan bağımsız, tek başına da kullanılabilir.
    Modelspace'deki her türlü HATCH entity'sini kaldırır.
    Blok tanımları içindekiler için clean_block_hatches() kullanın.

    Args:
        dxf_path   : Kaynak DXF dosyasının tam yolu
        output_path: Çıktı yolu (boş bırakılırsa _HATCH suffix eklenir)
    """
    import ezdxf as _ezdxf
    from pathlib import Path
    from collections import Counter as _Counter

    doc = _ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    hatches = [e for e in msp if e.dxftype() == "HATCH"]
    layer_dist = dict(_Counter(e.dxf.layer for e in hatches).most_common())

    for e in hatches:
        msp.delete_entity(e)

    if not output_path:
        p = Path(dxf_path)
        output_path = str(p.parent / (p.stem + "_HATCH" + p.suffix))

    doc.saveas(output_path)

    result = {
        "hatch_silindi": len(hatches),
        "layer_dagilimi": layer_dist,
        "cikti_dosya": output_path,
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def delete_tefris(dxf_path: str, output_path: str = "") -> str:
    """
    ADIM 5 — Tefris/mobilya temizleme.
    TEFRIS, _AB_TEFRIS, TEFRIS_YATAK, AKS_TEFRIS gibi
    döşeme/mobilya layer'larındaki tüm entity'leri kaldırır.
    Duvar iskeletini daha net ortaya çıkarır.

    Args:
        dxf_path   : Kaynak DXF (genellikle clean_hatch çıktısı)
        output_path: Çıktı yolu (boş bırakılırsa _TEFRIS suffix eklenir)
    """
    import ezdxf as _ezdxf
    from pathlib import Path
    from collections import Counter as _Counter

    def _is_tefris_layer(lu: str) -> bool:
        return ("TEFR" in lu or "MOBIL" in lu or "FURNITURE" in lu
                or "AKS_TEFR" in lu or "AKS-TEFR" in lu)

    doc = _ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    to_del = [e for e in msp if _is_tefris_layer(e.dxf.layer.upper()) and not _is_protected(e.dxf.layer)]
    layer_dist = dict(_Counter(e.dxf.layer for e in to_del).most_common())

    for e in to_del:
        msp.delete_entity(e)

    if not output_path:
        p = Path(dxf_path)
        output_path = str(p.parent / (p.stem + "_TEFRIS" + p.suffix))

    doc.saveas(output_path)

    result = {
        "tefris_silindi": len(to_del),
        "layer_dagilimi": layer_dist,
        "cikti_dosya": output_path,
        "sonraki_adim": "Alan tespiti için detect_rooms() kullanın"
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def delete_ceiling(dxf_path: str, output_path: str = "") -> str:
    """
    ADIM 6 — Tavan/asma tavan temizleme.
    ASMA TAVAN, TAVAN, CEILING gibi layer'lardaki
    tüm entity'leri kaldırır.

    Args:
        dxf_path   : Kaynak DXF (genellikle delete_tefris çıktısı)
        output_path: Çıktı yolu (boş bırakılırsa _TAVAN suffix eklenir)
    """
    import ezdxf as _ezdxf
    from pathlib import Path
    from collections import Counter as _Counter

    def _is_ceiling_layer(lu: str) -> bool:
        return ("TAVAN" in lu or "ASMA" in lu or "CEILING" in lu
                or "SUSPEND" in lu)

    doc = _ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    to_del = [e for e in msp if _is_ceiling_layer(e.dxf.layer.upper()) and not _is_protected(e.dxf.layer)]
    layer_dist = dict(_Counter(e.dxf.layer for e in to_del).most_common())

    for e in to_del:
        msp.delete_entity(e)

    if not output_path:
        p = Path(dxf_path)
        output_path = str(p.parent / (p.stem + "_TAVAN" + p.suffix))

    doc.saveas(output_path)

    result = {
        "tavan_silindi": len(to_del),
        "layer_dagilimi": layer_dist,
        "cikti_dosya": output_path,
        "sonraki_adim": "detect_rooms() ile alan tespiti yapabilirsiniz"
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def delete_linye(dxf_path: str, output_path: str = "") -> str:
    """
    ADIM 7 — Elektrik kablo linye numaralarını temizleme.
    LİNYE layer'ındaki daire + çizgi + text üçlüsünden oluşan
    A1, A2...An linye numarası sembollerini kaldırır.

    Args:
        dxf_path   : Kaynak DXF (genellikle delete_ceiling çıktısı)
        output_path: Çıktı yolu (boş bırakılırsa _LINYE suffix eklenir)
    """
    import ezdxf as _ezdxf
    from pathlib import Path
    from collections import Counter as _Counter

    def _is_linye_layer(lu: str) -> bool:
        return "NYE" in lu or "LINYE" in lu or "L\ufffdNYE" in lu

    doc = _ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    to_del = [e for e in msp if _is_linye_layer(e.dxf.layer.upper()) and not _is_protected(e.dxf.layer)]
    layer_dist = dict(_Counter(e.dxf.layer for e in to_del).most_common())
    type_dist  = dict(_Counter(e.dxftype()  for e in to_del).most_common())

    for e in to_del:
        msp.delete_entity(e)

    if not output_path:
        p = Path(dxf_path)
        output_path = str(p.parent / (p.stem + "_LINYE" + p.suffix))

    doc.saveas(output_path)

    result = {
        "linye_silindi": len(to_del),
        "layer_dagilimi": layer_dist,
        "tip_dagilimi": type_dist,
        "cikti_dosya": output_path,
        "sonraki_adim": "delete_electric_component() ile elektrik bileşenlerini temizleyin"
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def delete_electric_component(dxf_path: str, output_path: str = "") -> str:
    """
    ADIM 8 — Elektrik bileşen/sembol temizleme.
    Kablo sayı göstergeleri ('3','4','b','kom','ADİ','EKOM'),
    priz, anahtar, sigorta, panel sembollerini kaldırır.

    Hedef layer'lar: *ELEKTR*, *SEMBOL*, *PRIZ*, *ANAHTAR*,
                     *SIGORTA*, *PANEL*, *ROZET*, *SWITCH*, *SOCKET*
    Hedef bloklar  : '2','3','4','5','b','kom','ADİ','EKOM','EADI',
                     ve *HAT* layer'larındaki tüm INSERT'ler

    Args:
        dxf_path   : Kaynak DXF
        output_path: Çıktı yolu (boş bırakılırsa _ELEC suffix eklenir)
    """
    import ezdxf as _ezdxf
    from pathlib import Path
    from collections import Counter as _Counter

    # Kablo sayı/tip gösterge blok isimleri
    _CABLE_MARKER_BLOCKS = {
        '1','2','3','4','5','6','7','8','9',
        'b','B','kom','KOM','ADİ','ADI','EKOM',
        'EADI','EADİ','EVAV','VAV','IT'
    }

    # Tüm entity tiplerini silecek elektrik layerları
    _FULL_DELETE_LAYERS = {
        'ELKMETIN', 'ELK.YAZI', 'ELK YAZI', 'DEHA_ELK',
        'DEHA_ELK YAZI', 'E-SEMBOL', 'ELKSEMBOL',
        'MSE01', 'MCM01', 'DASH1L', 'S_M',
        '+IYI AYDINLATMA', 'AYDINLATMAHAT', 'E-AYDINLATMA',
    }

    def _is_electric_layer(lu: str) -> bool:
        if lu in _FULL_DELETE_LAYERS:
            return True
        return any(k in lu for k in (
            'ELEKTR', 'SEMBOL', 'PRIZ', 'ANAHTAR',
            'SIGORTA', 'PANEL', 'ROZET', 'SWITCH',
            'SOCKET', 'AYDINLATMA', 'ARMATUR', 'ARMATÜR',
            'ELK', 'ANAHTARPRIZ',
        )) or 'HAT' in lu

    # Tüm entity tipleri silinecek layer'lar (sadece INSERT değil)
    _ALL_TYPES = ('INSERT', 'TEXT', 'MTEXT', 'LINE', 'ARC', 'CIRCLE', 'LWPOLYLINE')

    doc = _ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    to_del = []
    for e in msp:
        if _is_protected(e.dxf.layer):
            continue
        layer = e.dxf.layer.upper()
        etype = e.dxftype()
        # INSERT: elektrik layer veya kablo marker block adı
        if etype == 'INSERT':
            name = e.dxf.name
            if _is_electric_layer(layer) or name in _CABLE_MARKER_BLOCKS:
                to_del.append(e)
        # Diğer tipler: sadece açıkça elektrik olan layer'larda
        elif etype in _ALL_TYPES and _is_electric_layer(layer):
            to_del.append(e)

    layer_dist = dict(_Counter(e.dxf.layer for e in to_del).most_common())
    block_dist = dict(_Counter(e.dxf.name  for e in to_del
                               if e.dxftype() == 'INSERT').most_common())

    for e in to_del:
        msp.delete_entity(e)

    # TARAMA layer LWPOLYLINE (kablo numarası çerçeveleri)
    tarama_del = [e for e in msp
                  if e.dxf.layer.upper() in ('TARAMA', 'TAR1', 'TAR2', 'TAR3')
                  and e.dxftype() in ('LWPOLYLINE', 'LINE', 'HATCH')]
    for e in tarama_del:
        msp.delete_entity(e)

    # Blok tanımları içindeki elektrik layer entity'lerini temizle
    # (Xref flatten blokları: MSE01, MCM01, DASH1L vb. içerebilir)
    block_internal_del = 0
    for block in doc.blocks:
        if block.name.startswith('*'):
            continue
        elec_in_block = [e for e in block
                         if _is_electric_layer(e.dxf.layer.upper())
                         and e.dxftype() in _ALL_TYPES]
        for e in elec_in_block:
            try:
                block.delete_entity(e)
                block_internal_del += 1
            except Exception:
                pass

    if not output_path:
        p = Path(dxf_path)
        output_path = str(p.parent / (p.stem + "_ELEC" + p.suffix))

    doc.saveas(output_path)

    result = {
        "silinen_insert": len(to_del),
        "tarama_silindi": len(tarama_del),
        "blok_ici_silindi": block_internal_del,
        "layer_dagilimi": layer_dist,
        "blok_dagilimi":  block_dist,
        "cikti_dosya":    output_path,
        "sonraki_adim":   "detect_rooms() ile alan tespiti yapabilirsiniz"
    }
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def delete_bara(dxf_path: str, output_path: str = "") -> str:
    """
    ADIM 9 — Bara (busbar) ve bağlı kablo çizgilerini temizleme.
    Elektrik dağıtım baralarını ve bağlantı kablolarını kaldırır.

    İki yöntemle tespit eder:
      A) Layer adı: BARA, BUS, PANO, PANEL, DAGIT, OG, AG HAT içeren layer'lar
      B) Renk kodu: Layer='0' üzerinde explicit color=3 (yeşil) olan LINE/LWPOLYLINE/ARC
         (elektrikçiler default layer üzerine yeşil renk ile bara çizer)

    Kaldırılanlar:
      • Bara layer'larındaki tüm entity'ler
      • Layer '0' + explicit green (color=3) LINE / LWPOLYLINE / ARC
      • AydPrizLinyesi ve benzeri panel linyesi layer'ları

    Args:
        dxf_path   : Kaynak DXF (genellikle delete_electric_component çıktısı)
        output_path: Çıktı yolu (boş bırakılırsa _BARA suffix eklenir)
    """
    import ezdxf as _ezdxf
    from pathlib import Path
    from collections import Counter as _Counter

    _BARA_LAYER_KWS = (
        'BARA', 'BUS', 'PANO', 'PANEL', 'DAGIT', 'DAĞIT',
        'OG HAT', 'AG HAT', 'TRAFO', 'KESICI', 'FIDERI',
        'AYDINLATMAPRIZ', 'AYDPRIZ', 'LINYE',
    )
    _WIRE_TYPES = ('LINE', 'LWPOLYLINE', 'ARC', 'SPLINE')

    def _is_bara_layer(lu: str) -> bool:
        return any(k in lu for k in _BARA_LAYER_KWS)

    doc = _ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    to_del = []
    for e in msp:
        if _is_protected(e.dxf.layer):
            continue
        layer = e.dxf.layer.upper()
        etype = e.dxftype()

        # A) Bilinen bara layer adları — tüm entity tipleri
        if _is_bara_layer(layer):
            to_del.append(e)
            continue

        # B) Layer='0' + explicit color=3 (yeşil) — INSERT hariç
        if e.dxf.layer == '0' and etype in _WIRE_TYPES:
            try:
                if e.dxf.color == 3:
                    to_del.append(e)
            except Exception:
                pass

    layer_dist = dict(_Counter(e.dxf.layer for e in to_del).most_common())
    type_dist  = dict(_Counter(e.dxftype()  for e in to_del).most_common())

    for e in to_del:
        msp.delete_entity(e)

    if not output_path:
        p = Path(dxf_path)
        output_path = str(p.parent / (p.stem + "_BARA" + p.suffix))

    doc.saveas(output_path)

    result = {
        "bara_silindi": len(to_del),
        "layer_dagilimi": layer_dist,
        "tip_dagilimi": type_dist,
        "cikti_dosya": output_path,
        "sonraki_adim": "colorize_mahal_blocks() ile oda renklendirme yapabilirsiniz"
    }
    return __import__("json").dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def delete_mahal_markers(dxf_path: str, output_path: str = "") -> str:
    """
    ADIM 10 — Önceki colorizer çalışmasından kalan büyük işaret/yazıları sil.
    colorize_rooms_in_cad veya colorize_mahal_blocks tarafından DXF'e
    yazılmış ve sonraki çalışmalarda sorun çıkaran artifact'ları temizler.

    Kaldırılanlar:
      • MAHAL-KIRMIZI layer  (kırmızı daire + büyük "(TANIMSIZ)" yazıları)
      • MAHAL-YESIL  layer  (yeşil hatch/daire - eğer varsa)
      • MAHAL-MAVI   layer  (mavi hatch/daire - eğer varsa)
      • MAHAL-TANIMLI / MAHAL-TANIMSIZ layer'ları (eski format)
      • "(TANIMSIZ)" veya "(TANIMLI)" içeren tüm TEXT/MTEXT

    Args:
        dxf_path   : Kaynak DXF
        output_path: Çıktı yolu (boş bırakılırsa _CLEAN suffix eklenir)
    """
    import ezdxf as _ezdxf
    from pathlib import Path
    from collections import Counter as _Counter

    _MARKER_LAYERS = {
        'MAHAL-KIRMIZI', 'MAHAL-YESIL', 'MAHAL-MAVI',
        'MAHAL-TANIMLI', 'MAHAL-TANIMSIZ',
        'MAHAL-GREEN', 'MAHAL-RED', 'MAHAL-BLUE',
    }

    doc = _ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    to_del = []
    for e in msp:
        layer = e.dxf.layer.upper()
        # Bilinen marker layer'ları — tüm entity'ler
        if layer in {l.upper() for l in _MARKER_LAYERS}:
            to_del.append(e)
            continue
        # "(TANIMSIZ)" veya "(TANIMLI)" içeren büyük yazılar
        if e.dxftype() in ('TEXT', 'MTEXT'):
            try:
                txt = (e.dxf.text if e.dxftype() == 'TEXT'
                       else e.text if hasattr(e, 'text') else '')
                if 'TANIMSIZ' in txt.upper() or 'TANIMLI' in txt.upper():
                    to_del.append(e)
            except Exception:
                pass

    layer_dist = dict(_Counter(e.dxf.layer for e in to_del).most_common())
    type_dist  = dict(_Counter(e.dxftype()  for e in to_del).most_common())

    for e in to_del:
        msp.delete_entity(e)

    if not output_path:
        p = Path(dxf_path)
        output_path = str(p.parent / (p.stem + "_CLEAN" + p.suffix))

    doc.saveas(output_path)

    return __import__("json").dumps({
        "silindi": len(to_del),
        "layer_dagilimi": layer_dist,
        "tip_dagilimi": type_dist,
        "cikti_dosya": output_path,
        "sonraki_adim": "colorize_mahal_blocks() ile oda renklendirme yapabilirsiniz"
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def colorize_mahal_blocks(dxf_path: str) -> str:
    """
    ADIM 9 — MAHAL bloklarını GstarCAD'de renkli daire+metin olarak işaretle.
    DXF'teki MAHAL_MEVCUT_R01 (ve benzeri MAHAL*) INSERT bloklarını okur,
    GstarCAD'deki aktif çizimde her oda için renkli daire + isim çizer.

    Renk şeması:
      Yeşil  (MAHAL-YESIL,    color 3) → MAHAL adı tanımlı oda
      Mavi   (MAHAL-MAVI,     color 5) → Adı olmayan oda (koordinat var)
      Kırmızı(MAHAL-KIRMIZI,  color 1) → Alan bilgisi 0 olan oda

    GstarCAD açık ve aynı DXF yüklü olmalıdır.

    Args:
        dxf_path: DXF dosyasının tam yolu
    """
    import win32com.client
    import pythoncom
    import math as _math
    import time as _time
    import re as _re
    import ezdxf as _ezdxf

    LAYER_GREEN  = "MAHAL-YESIL"
    LAYER_BLUE   = "MAHAL-MAVI"
    LAYER_RED    = "MAHAL-KIRMIZI"

    def _parse_area(s: str) -> float:
        s = str(s).replace(",", ".").replace("m2", "").replace("m\u00b2", "").strip()
        m = _re.search(r"[\d.]+", s)
        try:
            return float(m.group()) if m else 0.0
        except ValueError:
            return 0.0

    def _find_mu(attrs: dict) -> str:
        for key in ("\u00dc", "M\u00fc", "MU", "M2", "M\ufffd"):
            if key in attrs and attrs[key].strip():
                return attrs[key]
        for v in attrs.values():
            v = v.strip()
            if v and ("m2" in v.lower() or "m\u00b2" in v):
                return v
        for k, v in attrs.items():
            if k.startswith("M") and len(k) <= 3 and v.strip():
                return v
        return ""

    # ── DXF'ten MAHAL bloklarını oku ────────────────────────────────────────
    doc = _ezdxf.readfile(dxf_path)
    msp_dxf = doc.modelspace()
    rooms = []
    for ent in msp_dxf:
        if ent.dxftype() != "INSERT":
            continue
        bn    = ent.dxf.name.upper()
        layer = ent.dxf.layer.upper()
        if "MAHAL" not in bn and "MAHAL" not in layer:
            continue
        if not hasattr(ent, "attribs") or not ent.attribs:
            continue
        attrs  = {a.dxf.tag.upper(): a.dxf.text for a in ent.attribs}
        name   = (attrs.get("ROOMOBJECTS:NAME") or attrs.get("NAME") or
                  attrs.get("MAHAL_ADI") or attrs.get("MAHAL") or "").strip()
        number = (attrs.get("MAHALNO") or attrs.get("ROOMOBJECTS:NUMBER") or "").strip()
        area_s = _find_mu(attrs)
        area   = _parse_area(area_s)
        rooms.append({
            "name": name, "number": number, "area": area,
            "x": ent.dxf.insert.x, "y": ent.dxf.insert.y,
        })

    if not rooms:
        return '{"error": "MAHAL blogu bulunamadi. DXF dosyasini kontrol edin."}'

    # ── GstarCAD bağlantısı ──────────────────────────────────────────────────
    for attempt in range(3):
        try:
            acad = win32com.client.GetActiveObject("GstarCAD.Application")
            # Açık belge yoksa DXF'i otomatik aç
            try:
                cad_doc = acad.ActiveDocument
                _ = cad_doc.ModelSpace.Count
            except Exception:
                cad_doc = acad.Documents.Open(dxf_path)
                _time.sleep(2)
            cad_msp = cad_doc.ModelSpace
            break
        except Exception:
            if attempt == 2:
                raise
            _time.sleep(3)

    # Önceki MAHAL katmanlarını temizle
    for layer_name in (LAYER_GREEN, LAYER_BLUE, LAYER_RED):
        try:
            cmd = f'(command "._ERASE" (ssget "X" (list (cons 8 "{layer_name}"))) "")\n'
            cad_doc.SendCommand(cmd)
        except Exception:
            pass

    # Katmanları oluştur
    for name, color in ((LAYER_GREEN, 3), (LAYER_BLUE, 5), (LAYER_RED, 1)):
        try:
            lyr = cad_doc.Layers.Add(name)
        except Exception:
            lyr = cad_doc.Layers.Item(name)
        lyr.Color = color

    green = blue = red = 0

    for room in rooms:
        dx, dy  = room["x"], room["y"]
        area    = room["area"]
        name    = room["name"]
        number  = room["number"]

        # Daire yarıçapı: alanı temsil eden daire (mm²)
        if area > 0:
            r = _math.sqrt(area * 1e6 / _math.pi)   # m² → mm²
        else:
            r = 500.0

        if name:
            layer = LAYER_GREEN; color = 3; green += 1
        elif area > 0:
            layer = LAYER_BLUE;  color = 5; blue  += 1
        else:
            layer = LAYER_RED;   color = 1; red   += 1

        try:
            pt = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8, [dx, dy, 0.0])
            circ = cad_msp.AddCircle(pt, r)
            circ.Layer = layer
            circ.Color = color

            label = f"{name} ({number})" if number else name or number or "?"
            if area > 0:
                label += f" {area:.1f}m2"
            tp = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8, [dx, dy + r * 0.1, 0.0])
            txt = cad_msp.AddText(label, tp, r * 0.15)
            txt.Layer = layer
            txt.Color = color
        except Exception:
            pass

    cad_doc.Regen(1)

    result = {
        "toplam_oda": len(rooms),
        "yesil_isimli": green,
        "mavi_isimsiz": blue,
        "kirmizi_alansiz": red,
        "summary": (
            f"{green} yesil (isimli), {blue} mavi (isimsiz), "
            f"{red} kirmizi (alansiz) — toplam {len(rooms)} oda"
        )
    }
    return __import__("json").dumps(result, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    mcp.run()
