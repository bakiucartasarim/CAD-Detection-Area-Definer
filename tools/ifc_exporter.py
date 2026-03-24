"""
ifc_exporter.py
Detected wall polygons → IFC export.

Her kapalı duvar polyline'ı:
  - IfcSpace  : odanın iç alanı (polygon, 3m yükseklik)
  - IfcWall   : her kenar için 10cm kalınlık, 3m yükseklik

Koordinat:  unit_factor=1e-4 (cm) → linear_scale=0.01 → 1 DXF unit = 1 cm = 0.01 m
"""
from __future__ import annotations
import math
import uuid
import time
import ifcopenshell


def export_walls_to_ifc(
    dxf_path: str,
    output_path: str,
    wall_height_m: float = 3.0,
    wall_thickness_m: float = 0.10,
) -> dict:
    from parsers.dxf_parser import parse_dxf
    from parsers.element_classifier import classify_all_layers

    data = parse_dxf(dxf_path)
    layer_types = classify_all_layers(data["layers"])
    uf = data["unit_factor"]
    linear_scale = math.sqrt(uf)   # 1e-4 → 0.01 m/unit

    wall_layers = {n for n, t in layer_types.items() if t == "walls"}

    # ── Kapalı duvar polyline'larını topla ─────────────────────────────────
    rooms: list[dict] = []
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
        pts_m = [[p[0] * linear_scale, p[1] * linear_scale] for p in pts]
        cx = sum(p[0] for p in pts_m) / len(pts_m)
        cy = sum(p[1] for p in pts_m) / len(pts_m)
        rooms.append({"pts": pts_m, "cx": cx, "cy": cy,
                      "name": "", "number": "", "area_attr": ""})

    # ── MAHAL BLOCK INSERT'lerinden oda isimlerini oku ─────────────────────
    labels = _collect_labels(data["entities"], linear_scale)

    # Her odaya en yakın etiketi eşleştir
    used: set[int] = set()
    for room in rooms:
        best_i, best_d = -1, float("inf")
        for i, lbl in enumerate(labels):
            if i in used:
                continue
            d = math.hypot(lbl["x"] - room["cx"], lbl["y"] - room["cy"])
            if d < best_d:
                best_d, best_i = d, i
        if best_i >= 0 and best_d < 50.0:   # 50 m eşik (koordinatlar m cinsinden)
            room["name"]      = labels[best_i]["name"]
            room["number"]    = labels[best_i]["number"]
            room["area_attr"] = labels[best_i]["area"]
            used.add(best_i)

    ifc = _build_ifc(rooms, wall_height_m, wall_thickness_m)
    ifc.write(output_path)

    named = sum(1 for r in rooms if r["name"])
    return {
        "spaces": len(rooms),
        "named_spaces": named,
        "walls": sum(len(r["pts"]) for r in rooms),
        "slabs": len(rooms) * 2,   # zemin + tavan her oda için
        "output": output_path,
    }


def _collect_labels(entities: list, linear_scale: float) -> list[dict]:
    """MAHAL BLOCK INSERT'lerinden isim/numara/alan oku."""
    labels = []
    for ent in entities:
        if ent["type"] != "INSERT":
            continue
        layer = ent.get("layer", "").lower()
        if not any(kw in layer for kw in ("mahal", "room", "space")):
            continue
        attrs = ent.get("attrs", {})
        if not attrs:
            continue
        name   = (attrs.get("ROOMOBJECTS:NAME") or attrs.get("NAME") or
                  attrs.get("MAHAL_ADI") or "").strip()
        number = (attrs.get("ROOMOBJECTS:NUMBER") or attrs.get("NUMBER") or "").strip()
        area   = (attrs.get("ALAN:NAME") or attrs.get("ALAN") or
                  attrs.get("AREA") or "").strip()
        if not name and not number:
            continue
        labels.append({
            "x":      ent.get("x", 0.0) * linear_scale,
            "y":      ent.get("y", 0.0) * linear_scale,
            "name":   name,
            "number": number,
            "area":   area,
        })
    return labels


# ── helpers ───────────────────────────────────────────────────────────────────

def _uid():
    return ifcopenshell.guid.compress(uuid.uuid4().hex)


def _cp3(ifc, x, y, z=0.0):
    return ifc.createIfcCartesianPoint([float(x), float(y), float(z)])


def _cp2(ifc, x, y):
    return ifc.createIfcCartesianPoint([float(x), float(y)])


def _dir3(ifc, x, y, z):
    return ifc.createIfcDirection([float(x), float(y), float(z)])


def _ax3(ifc, origin, z_dir=None, x_dir=None):
    return ifc.createIfcAxis2Placement3D(
        origin,
        z_dir or _dir3(ifc, 0, 0, 1),
        x_dir or _dir3(ifc, 1, 0, 0),
    )


def _placement(ifc, parent, x=0., y=0., z=0., angle=0.):
    return ifc.createIfcLocalPlacement(
        parent,
        _ax3(ifc,
             _cp3(ifc, x, y, z),
             _dir3(ifc, 0, 0, 1),
             _dir3(ifc, math.cos(angle), math.sin(angle), 0)),
    )


# ── main builder ──────────────────────────────────────────────────────────────

def _build_ifc(rooms, wall_h, wall_t):
    ifc = ifcopenshell.file(schema="IFC4")

    org  = ifc.createIfcOrganization(None, "CAD Detection", None)
    pers = ifc.createIfcPerson(None, "Detector", "CAD")
    pao  = ifc.createIfcPersonAndOrganization(pers, org)
    app  = ifc.createIfcApplication(org, "1.0", "CAD Detection Area Definer", "CDAD")
    owner = ifc.createIfcOwnerHistory(pao, app, None, "ADDED", None, pao, app, int(time.time()))

    ctx = ifc.createIfcGeometricRepresentationContext(
        None, "Model", 3, 1.0e-5,
        _ax3(ifc, _cp3(ifc, 0, 0, 0)),
        None,
    )
    body = ifc.createIfcGeometricRepresentationSubContext(
        "Body", "Model", None, None, None, None, ctx, None, "MODEL_VIEW", None,
    )

    units = ifc.createIfcUnitAssignment([
        ifc.createIfcSIUnit(None, "LENGTHUNIT",      None, "METRE"),
        ifc.createIfcSIUnit(None, "AREAUNIT",        None, "SQUARE_METRE"),
        ifc.createIfcSIUnit(None, "VOLUMEUNIT",      None, "CUBIC_METRE"),
        ifc.createIfcSIUnit(None, "PLANEANGLEUNIT",  None, "RADIAN"),
    ])

    project  = ifc.createIfcProject(_uid(), owner, "CAD Walls Project", None,
                                     None, None, None, None, units)
    site     = ifc.createIfcSite    (_uid(), owner, "Site",        None, None,
                                     _placement(ifc, None), None, None, "ELEMENT", None)
    building = ifc.createIfcBuilding(_uid(), owner, "Building",    None, None,
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
    slabs  = []
    SLAB_T = 0.20   # 20 cm döşeme kalınlığı

    for idx, room in enumerate(rooms):
        pts    = room["pts"]
        r_name = room.get("name") or f"Oda {idx + 1}"
        r_num  = room.get("number") or str(idx + 1)
        area   = _polygon_area(pts)

        # IfcSpace
        poly_pts = [_cp2(ifc, p[0], p[1]) for p in pts]
        poly_pts.append(poly_pts[0])
        profile = ifc.createIfcArbitraryClosedProfileDef(
            "AREA", None, ifc.createIfcPolyline(poly_pts))
        solid = ifc.createIfcExtrudedAreaSolid(
            profile, _ax3(ifc, _cp3(ifc, 0, 0, 0)), _dir3(ifc, 0, 0, 1), float(wall_h))
        shape = ifc.createIfcProductDefinitionShape(None, None, [
            ifc.createIfcShapeRepresentation(body, "Body", "SweptSolid", [solid])])
        space = ifc.createIfcSpace(
            _uid(), owner,
            r_name,           # Name  = oda adı (MUAYENE ODASI vb.)
            r_num,            # Description = oda numarası (Z_02 vb.)
            None,
            _placement(ifc, storey.ObjectPlacement), shape, None, "ELEMENT", "INTERNAL")
        _pset(ifc, owner, space, "Pset_SpaceCommon", {
            "NetFloorArea": round(area, 3),
            "SpaceID":      idx + 1,
            "MahalAdi":     r_name,
            "MahalNo":      r_num,
        })
        spaces.append(space)

        # ── IfcSlab: zemin (z=0) ve tavan (z=wall_h) ─────────────────────
        for slab_z, slab_label in [(0.0, "Zemin"), (wall_h, "Tavan")]:
            s_poly = [_cp2(ifc, p[0], p[1]) for p in pts]
            s_poly.append(s_poly[0])
            s_profile = ifc.createIfcArbitraryClosedProfileDef(
                "AREA", None, ifc.createIfcPolyline(s_poly))
            s_solid = ifc.createIfcExtrudedAreaSolid(
                s_profile, _ax3(ifc, _cp3(ifc, 0, 0, 0)),
                _dir3(ifc, 0, 0, 1), float(SLAB_T))
            s_shape = ifc.createIfcProductDefinitionShape(None, None, [
                ifc.createIfcShapeRepresentation(body, "Body", "SweptSolid", [s_solid])])
            slab = ifc.createIfcSlab(
                _uid(), owner,
                f"{slab_label}_{r_name or idx+1}", None, None,
                _placement(ifc, storey.ObjectPlacement, 0., 0., slab_z),
                s_shape, None,
                "FLOOR" if slab_label == "Zemin" else "ROOF")
            slabs.append(slab)

        # IfcWall per edge
        n = len(pts)
        for j in range(n):
            p1 = pts[j]
            p2 = pts[(j + 1) % n]
            dx = p2[0] - p1[0]
            dy = p2[1] - p1[1]
            length = math.hypot(dx, dy)
            if length < 0.01:
                continue
            angle = math.atan2(dy, dx)
            w_profile = ifc.createIfcRectangleProfileDef(
                "AREA", None,
                ifc.createIfcAxis2Placement2D(
                    _cp2(ifc, length / 2, wall_t / 2), None),
                float(length), float(wall_t))
            w_solid = ifc.createIfcExtrudedAreaSolid(
                w_profile, _ax3(ifc, _cp3(ifc, 0, 0, 0)),
                _dir3(ifc, 0, 0, 1), float(wall_h))
            w_shape = ifc.createIfcProductDefinitionShape(None, None, [
                ifc.createIfcShapeRepresentation(body, "Body", "SweptSolid", [w_solid])])
            wall = ifc.createIfcWall(
                _uid(), owner, f"Wall_R{idx}_S{j}", None, None,
                _placement(ifc, storey.ObjectPlacement,
                           p1[0], p1[1], 0., angle),
                w_shape, None, "SOLIDWALL")
            walls.append(wall)

    if spaces:
        ifc.createIfcRelContainedInSpatialStructure(
            _uid(), owner, None, None, spaces, storey)
    if walls:
        ifc.createIfcRelContainedInSpatialStructure(
            _uid(), owner, None, None, walls, storey)
    if slabs:
        ifc.createIfcRelContainedInSpatialStructure(
            _uid(), owner, None, None, slabs, storey)

    return ifc


def _pset(ifc, owner, element, name, props):
    ifc_props = []
    for k, v in props.items():
        if isinstance(v, float):
            val = ifc.createIfcReal(v)
        elif isinstance(v, int):
            val = ifc.createIfcInteger(v)
        else:
            val = ifc.createIfcLabel(str(v))
        ifc_props.append(ifc.createIfcPropertySingleValue(k, None, val, None))
    pset = ifc.createIfcPropertySet(_uid(), owner, name, None, ifc_props)
    ifc.createIfcRelDefinesByProperties(_uid(), owner, None, None, [element], pset)


def _polygon_area(pts):
    n = len(pts)
    a = 0.0
    for i in range(n):
        j = (i + 1) % n
        a += pts[i][0] * pts[j][1] - pts[j][0] * pts[i][1]
    return abs(a) / 2.0
