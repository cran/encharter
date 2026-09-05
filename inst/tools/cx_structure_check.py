#!/usr/bin/env python3
# Structural validator for chartex (cx:) chart parts, driven by the Open XML
# SDK schema data - the same schema definitions OOXMLValidator's SDK-based
# validation is generated from. Checks element names, child sequence order,
# attribute sets, and enum values. The SDK json is downloaded on first use
# and cached next to this script.
#
# Usage: python3 cx_structure_check.py chart1.xml [chart2.xml ...]
# Exit code = number of failing files.
# Resolves child types through the parent's particle (context-aware), so
# name collisions like the two cx:visibility types are handled.
import json, sys, re
import xml.etree.ElementTree as ET

CX = "http://schemas.microsoft.com/office/drawing/2014/chartex"
A  = "http://schemas.openxmlformats.org/drawingml/2006/main"

SCHEMA_URL = ("https://raw.githubusercontent.com/dotnet/Open-XML-SDK/main/"
              "data/schemas/schemas_microsoft_com_office_drawing_2014_chartex.json")
import os, urllib.request
cache = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chartex_schema.json")
if not os.path.exists(cache):
    urllib.request.urlretrieve(SCHEMA_URL, cache)
d = json.load(open(cache))
types = {t["Name"]: t for t in d["Types"]}
enums = {e["Name"]: {f.get("Value") for f in e.get("Facets", [])} for e in d.get("Enums", [])}

def enum_of(attr):
    m = re.search(r"ChartDrawing\.(\w+)>", attr.get("Type", ""))
    return m.group(1) if m else None

def slots(particle):
    out = []
    if particle.get("Kind") == "Sequence":
        for it in particle.get("Items", []):
            m = {}
            collect(it, m)
            out.append(m)
    else:
        m = {}
        collect(particle, m)
        if m: out.append(m)
    return out

def collect(p, into):
    if "Name" in p:
        into[p["Name"].split("/")[-1]] = p["Name"]
    for it in p.get("Items", []):
        collect(it, into)

def qname(el):
    if el.tag.startswith("{" + CX + "}"): return "cx:" + el.tag.split("}")[1]
    if el.tag.startswith("{" + A + "}"):  return "a:"  + el.tag.split("}")[1]
    return el.tag

errors = []

def check(el, tref, path):
    name = qname(el)
    t = types.get(tref)
    if t is None:
        if name.startswith("cx:"):
            errors.append(f"{path}: no schema type for {name} ({tref})")
        return
    path = f"{path}/{name}"
    allowed = {a["QName"].lstrip(":"): a for a in t.get("Attributes", [])}
    # attributes defined on SDK base classes, not listed per-type in the json
    if "CT_TickMarks/" in tref: allowed.setdefault("type", {"Type": "EnumValue<...ChartDrawing.TickMarksType>"})
    for k, v in el.attrib.items():
        k = k.split("}")[-1]
        if k not in allowed:
            # Known deviation: Excel writes <cx:binSize val=".."/> / binCount
            # although the schema types them as simple content. Accepted.
            if k == "val" and ("cx:binSize" in tref or "cx:binCount" in tref):
                continue
            errors.append(f"{path}: attribute '{k}' not allowed on {tref}")
            continue
        en = enum_of(allowed[k])
        if en and en in enums and v not in enums[en]:
            errors.append(f"{path}: @{k}='{v}' not in {en} {sorted(enums[en])}")
    part = t.get("Particle") or {}
    sl = slots(part)
    last = 0
    for c in el:
        cn = qname(c)
        if not cn.startswith("cx:"):
            continue  # a:* content validated via the ECMA schemas elsewhere
        idx = next((i for i in range(len(sl)) if cn in sl[i]), None)
        if idx is None:
            errors.append(f"{path}: unexpected child {cn} on {tref}")
            continue
        if idx < last:
            errors.append(f"{path}: child {cn} out of sequence on {tref}")
        last = max(last, idx)
        check(c, sl[idx][cn], path)

fail = 0
for f in sys.argv[1:]:
    errors.clear()
    root = ET.parse(f).getroot()
    check(root, "cx:CT_ChartSpace/cx:chartSpace", "")
    print(f"{f.split('/')[-1]}: {'OK' if not errors else 'FAIL'}")
    for e in errors[:8]: print("   ", e)
    fail += bool(errors)
sys.exit(fail)
