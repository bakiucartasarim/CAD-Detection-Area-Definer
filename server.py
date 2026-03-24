"""
server.py — CAD Detection Area Definer MCP Server

Claude'a şu araçları açar:
  • analyze_cad       → DXF dosyasındaki katmanları ve entity sayılarını döner
  • detect_rooms      → Oda poligonlarını, alanlarını ve etiketlerini çıkarır
  • classify_elements → Katmanları duvar/kapı/pencere/mobilya olarak sınıflandırır
  • get_unknown_layers→ Tanımlanamayan katmanları listeler
  • train_layer       → Yeni katman → tip eşleştirmesi öğretir
  • get_room_geometry → Tek bir odanın tam geometrisini döner
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


if __name__ == "__main__":
    mcp.run()
