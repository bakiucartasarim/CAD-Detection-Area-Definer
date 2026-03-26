"""
server.py — CAD Detection Area Definer MCP Server

Claude'a şu araçları açar:
  • analyze_cad       → DXF dosyasındaki katmanları ve entity sayılarını döner
  • detect_rooms      → Oda poligonlarını, alanlarını ve etiketlerini çıkarır
  • classify_elements → Katmanları duvar/kapı/pencere/mobilya olarak sınıflandırır
  • get_unknown_layers→ Tanımlanamayan katmanları listeler
  • train_layer       → Yeni katman → tip eşleştirmesi öğretir
  • get_room_geometry → Tek bir odanın tam geometrisini döner
  • export_walls_ifc  → Duvar polyline'larını IfcSpace+IfcWall olarak dışa aktarır
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


if __name__ == "__main__":
    mcp.run()
