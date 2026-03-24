"""
element_classifier.py
Classifies DXF layers and entities into semantic element types:
wall, door, window, furniture, column, stair, room, text, electrical, unknown.

Uses layer_registry.json as the knowledge base.
New mappings can be added via train_layer().
"""
from __future__ import annotations
import json
import os

REGISTRY_PATH = os.path.join(os.path.dirname(__file__), "..", "training", "layer_registry.json")

_ELEMENT_TYPES = [
    "walls", "doors", "windows", "furniture",
    "columns", "stairs", "rooms", "dimensions",
    "text", "electrical",
]


def _load_registry() -> dict[str, list[str]]:
    with open(REGISTRY_PATH, encoding="utf-8") as f:
        data = json.load(f)
    return {k: v for k, v in data.items() if not k.startswith("_")}


def _save_registry(registry: dict) -> None:
    with open(REGISTRY_PATH, encoding="utf-8") as f:
        data = json.load(f)
    for k in _ELEMENT_TYPES:
        if k in registry:
            data[k] = registry[k]
    with open(REGISTRY_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _normalize(s: str) -> str:
    # Handle both proper Unicode and garbled encodings
    result = s.lower()
    # Turkish characters (Unicode)
    for src, dst in [("ı","i"),("İ","i"),("ğ","g"),("Ğ","g"),("ş","s"),("Ş","s"),
                     ("ü","u"),("Ü","u"),("ö","o"),("Ö","o"),("ç","c"),("Ç","c")]:
        result = result.replace(src, dst)
    # Strip garbled bytes (non-ASCII junk from bad DXF encoding)
    result = "".join(c if ord(c) < 128 else "_" for c in result)
    return result


def classify_layer(layer_name: str) -> str:
    """Return the element type for a given layer name."""
    registry = _load_registry()
    nl = _normalize(layer_name)
    for etype, keywords in registry.items():
        if any(_normalize(kw) in nl for kw in keywords):
            return etype
    return "unknown"


def classify_all_layers(layers: dict) -> dict[str, str]:
    """
    Classify all layers from parse_dxf() output.
    Returns {layer_name: element_type}
    """
    return {name: classify_layer(name) for name in layers}


def train_layer(layer_name: str, element_type: str) -> dict:
    """
    Teach the system that a layer belongs to an element type.
    Saves to layer_registry.json permanently.
    Returns {"status": "ok", "layer": ..., "type": ...}
    """
    if element_type not in _ELEMENT_TYPES:
        return {"status": "error", "message": f"Unknown type '{element_type}'. Valid: {_ELEMENT_TYPES}"}

    registry = _load_registry()
    keywords = registry.get(element_type, [])
    nl = _normalize(layer_name)

    if nl not in [_normalize(k) for k in keywords]:
        keywords.append(layer_name.lower())
        registry[element_type] = keywords
        _save_registry(registry)
        return {"status": "ok", "layer": layer_name, "type": element_type, "action": "added"}

    return {"status": "ok", "layer": layer_name, "type": element_type, "action": "already_exists"}


def get_unknown_layers(layers: dict) -> list[str]:
    """Return layer names that couldn't be classified."""
    return [name for name, etype in classify_all_layers(layers).items() if etype == "unknown"]
