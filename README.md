# CAD Detection Area Definer

An MCP (Model Context Protocol) server that analyzes DXF/CAD files to detect and define room geometries by recognizing walls, doors, windows, columns, and furniture.

## Features

- **Room Detection** — 4-strategy pipeline: MAHAL blocks → HATCH fills → wall polygonize → closed polylines
- **Element Classification** — Classifies layers into: walls, doors, windows, furniture, columns, stairs, rooms, text, electrical
- **Self-Learning** — `train_layer()` permanently adds new layer→type mappings to the knowledge base
- **Turkish DXF Support** — Handles encoding issues, Turkish character normalization

## MCP Tools

| Tool | Description |
|------|-------------|
| `analyze_cad(dxf_path)` | Parse DXF: layers, entity counts, bounding box, unit |
| `detect_rooms(dxf_path)` | Detect all rooms with area, centroid, polygon, nearby doors/windows |
| `get_room_geometry(dxf_path, room_id)` | Full polygon points for a specific room |
| `classify_elements(dxf_path)` | Classify all layers by element type |
| `get_unknown_layers(dxf_path)` | List unclassified layers |
| `train_layer(layer_name, element_type)` | Teach a new layer mapping (persisted) |

## Setup

```bash
pip install -r requirements.txt
```

## Claude Desktop Integration

Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "cad-detector": {
      "command": "python",
      "args": ["C:/path/to/CAD-Detection-Area-Definer/server.py"]
    }
  }
}
```

## Detection Priority

```
1. MAHAL BLOCK  → INSERT entities with room name/area attributes
2. HATCH        → Boundary polygons (≥5 results, median area ≥3 m²)
3. Polygonize   → Shapely polygonize from wall LINE/LWPOLYLINE entities
4. Closed PLW   → Closed LWPOLYLINE entities on wall layers
```

## Training

When a layer is unrecognized, use `train_layer()` to teach it:

```
Claude: get_unknown_layers("myplan.dxf")
→ ["ZEMIN-KAPLAMA", "XREF-DUVAR"]

Claude: train_layer("ZEMIN-KAPLAMA", "furniture")
Claude: train_layer("XREF-DUVAR", "walls")
```

New mappings are saved to `training/layer_registry.json` and applied to all future DXF files.

## Project Structure

```
CAD-Detection-Area-Definer/
├── server.py                    # MCP server entry point
├── parsers/
│   ├── dxf_parser.py            # ezdxf entity extraction
│   ├── element_classifier.py    # Layer classification engine
│   └── geometry_engine.py       # Shapely room polygon builder
├── training/
│   ├── layer_registry.json      # Learned layer→type mappings
│   ├── samples/                 # Training DXF files
│   └── annotations/             # Manual layer annotations
└── requirements.txt
```
