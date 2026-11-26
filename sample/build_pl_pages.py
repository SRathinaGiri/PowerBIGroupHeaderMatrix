"""Usage: py -3.12 build_pl_pages.py <path to .Report folder>

Adds (or rebuilds) two test pages for the Group Header Matrix visual in a PBIR report that
contains the hotel P&L test model (see build_pl_model.py):

  - "P&L Test":        slicers (Year, Month, Region, Property) + matrix with P&L rows
  - "P&L by Property": slicers (Year, Month) + matrix with Region > Property on columns

Formatting already set on an existing matrix of the same page is kept. Run while the
report is closed in Power BI Desktop (or use `pbi report reload`), otherwise Desktop
overwrites the files on its next save.
"""
import json
import secrets
import shutil
import sys
from pathlib import Path

report = Path(sys.argv[1])
pages_dir = report / "definition" / "pages"
VC_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.13.0/schema.json"
PAGE_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json"
VISUAL_TYPE = "groupHeaderMatrix19740818"
FACT = "PL Fact"


def new_id() -> str:
    return secrets.token_hex(10)


def lit(value: str) -> dict:
    return {"expr": {"Literal": {"Value": value}}}


def measure(name: str) -> dict:
    return {
        "field": {"Measure": {"Expression": {"SourceRef": {"Entity": FACT}}, "Property": name}},
        "queryRef": f"{FACT}.{name}",
        "nativeQueryRef": name,
    }


def level(entity: str, hierarchy: str, lvl: str) -> dict:
    return {
        "field": {"HierarchyLevel": {
            "Expression": {"Hierarchy": {"Expression": {"SourceRef": {"Entity": entity}}, "Hierarchy": hierarchy}},
            "Level": lvl}},
        "queryRef": f"{entity}.{hierarchy}.{lvl}",
        "nativeQueryRef": f"{hierarchy} {lvl}",
    }


def column(entity: str, col: str) -> dict:
    return {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": col}}


def slicer(entity: str, col: str, x: float, width: float, z: int, selected=None, single=False) -> dict:
    objects = {
        "data": [{"properties": {"mode": lit("'Dropdown'")}}],
        "selection": [{"properties": {"singleSelect": lit("true" if single else "false")}}],
    }
    if selected:
        objects["general"] = [{"properties": {"filter": {"filter": {
            "Version": 2,
            "From": [{"Name": "s", "Entity": entity, "Type": 0}],
            "Where": [{"Condition": {"In": {
                "Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": col}}],
                "Values": [[{"Literal": {"Value": v}}] for v in selected]
            }}}]
        }}}}]
    return {
        "$schema": VC_SCHEMA,
        "name": new_id(),
        "position": {"x": x, "y": 8, "z": z, "height": 56, "width": width, "tabOrder": z},
        "visual": {
            "visualType": "slicer",
            "query": {"queryState": {"Values": {"projections": [
                {"field": column(entity, col), "queryRef": f"{entity}.{col}", "nativeQueryRef": col, "active": True}
            ]}}},
            "objects": objects,
            "drillFilterOtherVisuals": True,
        },
    }


ROW_LEVELS = [level("PL Lines", "P&L", "Line"), level("PL Lines", "P&L", "Detail"), level("PL Lines", "P&L", "GL Code")]
STYLE_ROLES = {
    "rowBgColor": {"projections": [measure("Style Row Fill")]},
    "rowBold": {"projections": [measure("Style Row Bold")]},
    # Pairs by position with Values: 1st -> none, 2nd (Plan) -> yellow, the rest -> none
    "cellBgColor": {"projections": [measure("Style Fill None"), measure("Style Fill Plan"), measure("Style Fill None 2")]},
    "cellFontColor": {"projections": [measure("Style Font Below Plan")]},
}
DEFAULT_OBJECTS = {"rowFormatting": [{"properties": {"applyToRowHeaders": lit("true")}}]}


def matrix(query_state: dict, objects: dict, z: int) -> dict:
    return {
        "$schema": VC_SCHEMA,
        "name": new_id(),
        "position": {"x": 8, "y": 72, "z": z, "height": 640, "width": 1264, "tabOrder": z},
        "visual": {
            "visualType": VISUAL_TYPE,
            "query": {"queryState": query_state},
            "objects": objects,
            "drillFilterOtherVisuals": True,
        },
    }


pages_meta_path = pages_dir / "pages.json"
pages_meta = json.loads(pages_meta_path.read_text(encoding="utf-8"))


def remove_page(display_name: str):
    """Deletes a page by display name and returns the objects of its matrix (to keep formatting)."""
    kept = None
    for pid in list(pages_meta["pageOrder"]):
        pj = pages_dir / pid / "page.json"
        if not pj.exists() or json.loads(pj.read_text(encoding="utf-8")).get("displayName") != display_name:
            continue
        for vj in (pages_dir / pid / "visuals").glob("*/visual.json"):
            v = json.loads(vj.read_text(encoding="utf-8")).get("visual", {})
            if v.get("visualType") == VISUAL_TYPE:
                kept = v.get("objects")
        shutil.rmtree(pages_dir / pid)
        pages_meta["pageOrder"].remove(pid)
    return kept


def add_page(display_name: str, visuals: list) -> str:
    pid = new_id()
    (pages_dir / pid / "visuals").mkdir(parents=True)
    (pages_dir / pid / "page.json").write_text(json.dumps({
        "$schema": PAGE_SCHEMA, "name": pid, "displayName": display_name,
        "displayOption": "FitToPage", "height": 720, "width": 1280,
    }, indent=2), encoding="utf-8")
    for v in visuals:
        (pages_dir / pid / "visuals" / v["name"]).mkdir()
        (pages_dir / pid / "visuals" / v["name"] / "visual.json").write_text(json.dumps(v, indent=2), encoding="utf-8")
    pages_meta["pageOrder"].append(pid)
    return pid


# Page 1: P&L Test (lines down the side, measures across)
kept_test = remove_page("P&L Test")
test_values = ["Actual", "Plan", "Last Year", "vs Plan %", "vs LY %", "YTD Actual", "YTD Plan",
               "YTD Last Year", "YTD vs Plan %", "YTD vs LY %", "Full Year Plan"]
test_page = add_page("P&L Test", [
    slicer("PL Calendar", "Year", 8, 140, 100),
    slicer("PL Calendar", "Month", 156, 200, 200, selected=["'Aug 2026'"]),
    slicer("PL Property", "Region", 364, 180, 300),
    slicer("PL Property", "Property", 552, 260, 400),
    matrix({"row": {"projections": ROW_LEVELS}, "measure": {"projections": [measure(v) for v in test_values]}, **STYLE_ROLES},
           kept_test or DEFAULT_OBJECTS, 1000),
])

# Page 2: P&L by Property (Region > Property across, three measures each)
kept_prop = remove_page("P&L by Property")
prop_values = ["Actual", "Plan", "vs Plan %"]
add_page("P&L by Property", [
    slicer("PL Calendar", "Year", 8, 140, 100),
    slicer("PL Calendar", "Month", 156, 200, 200, selected=["'Aug 2026'"]),
    matrix({
        "row": {"projections": ROW_LEVELS},
        "columns": {"projections": [level("PL Property", "Location", "Region"), level("PL Property", "Location", "Property")]},
        "measure": {"projections": [measure(v) for v in prop_values]},
        **STYLE_ROLES,
    }, kept_prop or DEFAULT_OBJECTS, 1000),
])

pages_meta["activePageName"] = test_page
pages_meta_path.write_text(json.dumps(pages_meta, indent=2), encoding="utf-8")
print("pages:", ", ".join(json.loads((pages_dir / p / "page.json").read_text(encoding="utf-8"))["displayName"] for p in pages_meta["pageOrder"]))
