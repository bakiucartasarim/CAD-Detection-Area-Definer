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
        layer = e.dxf.layer.upper()
        if _is_cable_layer(layer) and e.dxftype() in ("LINE", "LWPOLYLINE", "INSERT", "MTEXT", "TEXT"):
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

    to_del = [e for e in msp if _is_tefris_layer(e.dxf.layer.upper())]
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

    to_del = [e for e in msp if _is_ceiling_layer(e.dxf.layer.upper())]
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

    to_del = [e for e in msp if _is_linye_layer(e.dxf.layer.upper())]
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
