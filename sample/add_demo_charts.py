"""Usage: py -3.12 add_demo_charts.py <path to .Report folder>

Adds two small charts beside the Group Header Matrix on the "Demo" page (Superstore data) to
show cross-filtering in both directions:
  - Total Sales by Region (bar)      -> Region is on the matrix columns
  - Total Sales by Category (donut)  -> Category is on the matrix rows
Clicking a bar/slice filters the matrix; clicking a matrix cell filters both charts.
Re-running replaces the two charts. Run while the report is closed in Power BI Desktop.
"""
import json
import shutil
import sys
from pathlib import Path

report = Path(sys.argv[1])
pages_dir = report / "definition" / "pages"
VC_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.13.0/schema.json"


def demo_page() -> Path:
    for pj in pages_dir.glob("*/page.json"):
        if json.loads(pj.read_text(encoding="utf-8")).get("displayName") == "Demo":
            return pj.parent
    raise SystemExit("Demo page not found")


def lit(value: str) -> dict:
    return {"expr": {"Literal": {"Value": value}}}


def chart(name: str, visual_type: str, entity: str, column: str, y: float, height: float, title: str) -> dict:
    return {
        "$schema": VC_SCHEMA,
        "name": name,
        "position": {"x": 930, "y": y, "z": 3000 + int(y), "height": height, "width": 340, "tabOrder": 3000 + int(y)},
        "visual": {
            "visualType": visual_type,
            "query": {"queryState": {
                "Category": {"projections": [{
                    "field": {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": column}},
                    "queryRef": f"{entity}.{column}", "nativeQueryRef": column, "active": True}]},
                "Y": {"projections": [{
                    "field": {"Measure": {"Expression": {"SourceRef": {"Entity": "OrderFact"}}, "Property": "Total Sales"}},
                    "queryRef": "OrderFact.Total Sales", "nativeQueryRef": "Total Sales"}]},
            }},
            "drillFilterOtherVisuals": True,
        },
        "visualContainerObjects": {
            "title": [{"properties": {"show": lit("true"), "text": lit(f"'{title}'")}}]
        },
    }


page = demo_page()
visuals = page / "visuals"
charts = [
    chart("demoRegionBars", "clusteredBarChart", "Places", "Region", 10, 330, "Total Sales by Region (click to filter)"),
    chart("demoCategoryDonut", "donutChart", "Categories", "Category", 350, 360, "Total Sales by Category (click to filter)"),
]
for c in charts:
    target = visuals / c["name"]
    if target.exists():
        shutil.rmtree(target)
    target.mkdir(parents=True)
    (target / "visual.json").write_text(json.dumps(c, indent=2), encoding="utf-8")
print("added to Demo page:", ", ".join(c["name"] for c in charts))
