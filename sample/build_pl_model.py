"""Usage: py -3.12 build_pl_model.py localhost:<port>   (port from `pbi connect`)

Builds a fictional hotel-group P&L test model in the open Power BI Desktop file via TOM,
for testing the Group Header Matrix visual (row formatting, dynamic format strings,
hierarchies, column groups).

Adds calculated tables PL Lines, PL Property, PL Calendar and PL Fact, relationships,
a "P&L" hierarchy and measures (several with dynamic format strings). Other tables are
not touched. Re-running replaces only these objects. All numbers are synthetic.
"""
import sys

sys.path.insert(0, r"C:\Users\acer\AppData\Roaming\Python\Python312\site-packages")
from pbi_cli.core.dotnet_loader import get_server_class  # noqa: E402

Server = get_server_class()
import Microsoft.AnalysisServices.Tabular as TOM  # noqa: E402

PORT = sys.argv[1] if len(sys.argv) > 1 else "localhost:50667"

# Line Key, Line, Detail, Line Order, Kind, Emphasis, monthly base amount (thousands)
LINES = r'''
DATATABLE(
    "Line Key", INTEGER, "Line", STRING, "Detail", STRING, "Line Order", INTEGER,
    "Kind", STRING, "Emphasis", BOOLEAN, "Base", DOUBLE,
    {
        { 11, "Room Revenue", "Leisure Stays", 1, "Value", FALSE, 820 },
        { 12, "Room Revenue", "Corporate Stays", 1, "Value", FALSE, 540 },
        { 21, "Food & Beverage", "Restaurants", 2, "Value", FALSE, 310 },
        { 22, "Food & Beverage", "Banquets & Events", 2, "Value", FALSE, 260 },
        { 31, "Discounts & Rebates", "Loyalty Rebates", 3, "Value", FALSE, -45 },
        { 32, "Discounts & Rebates", "Promotional Offers", 3, "Value", FALSE, -38 },
        { 40, "Total Revenue", "Total Revenue", 4, "Value", TRUE, 0 },
        { 51, "Staff Cost", "Front Office", 5, "Value", FALSE, -190 },
        { 52, "Staff Cost", "Housekeeping", 5, "Value", FALSE, -160 },
        { 53, "Staff Cost", "Kitchen", 5, "Value", FALSE, -120 },
        { 61, "Cost of F&B Sold", "Food Cost", 6, "Value", FALSE, -150 },
        { 62, "Cost of F&B Sold", "Beverage Cost", 6, "Value", FALSE, -55 },
        { 70, "Contribution", "Contribution", 7, "Value", TRUE, 0 },
        { 80, "Contribution Margin %", "Contribution Margin %", 8, "Ratio", FALSE, 0 },
        { 90, "Staff Cost % of Revenue", "Staff Cost % of Revenue", 9, "Ratio", FALSE, 0 },
        { 101, "Overheads", "Energy", 10, "Value", FALSE, -95 },
        { 102, "Overheads", "Maintenance", 10, "Value", FALSE, -70 },
        { 103, "Overheads", "Lease & Admin", 10, "Value", FALSE, -130 },
        { 110, "Operating Profit", "Operating Profit", 11, "Value", TRUE, 0 },
        { 120, "Operating Margin %", "Operating Margin %", 12, "Ratio", FALSE, 0 }
    }
)'''

PROPERTIES = r'''
DATATABLE(
    "Property", STRING, "Segment", STRING, "Weight", DOUBLE,
    { { "City Centre Hotel", "Urban", 0.45 }, { "Airport Inn", "Urban", 0.25 }, { "Beach Resort", "Leisure", 0.30 } }
)'''

# Calendar-year months Jan-2024 .. Dec-2026
CALENDAR = r'''
VAR Months = FILTER(CALENDAR(DATE(2024, 1, 1), DATE(2026, 12, 1)), DAY([Date]) = 1)
RETURN
    SELECTCOLUMNS(
        Months,
        "Month Start", [Date],
        "Month", FORMAT([Date], "MMM yyyy"),
        "Month No", YEAR([Date]) * 12 + MONTH([Date]),
        "Year", YEAR([Date]),
        "Quarter", "Q" & QUARTER([Date]),
        "Year Index", YEAR([Date]) - 2024
    )'''

FACT = r'''
VAR Months = SELECTCOLUMNS('PL Calendar', "Month Start", 'PL Calendar'[Month Start], "MN", 'PL Calendar'[Month No], "YI", 'PL Calendar'[Year Index])
VAR Lines = SELECTCOLUMNS(FILTER('PL Lines', 'PL Lines'[Base] <> 0), "Line Key", 'PL Lines'[Line Key], "B", 'PL Lines'[Base])
VAR Props = SELECTCOLUMNS('PL Property', "Property", 'PL Property'[Property], "W", 'PL Property'[Weight])
VAR Versions = DATATABLE("Version", STRING, { { "Actual" }, { "Plan" } })
VAR AllRows = GENERATE(GENERATE(GENERATE(Months, Lines), Props), Versions)
VAR Kept = FILTER(AllRows, [Version] = "Plan" || [Month Start] <= DATE(2026, 8, 1))
RETURN
    SELECTCOLUMNS(
        Kept,
        "Month Start", [Month Start],
        "Property", [Property],
        "Line Key", [Line Key],
        "Version", [Version],
        "Amount",
            ROUND(
                [B] * [W]
                    * (1 + 0.12 * SIN([MN] * 0.52 + [Line Key] / 7))
                    * POWER(1.09, [YI])
                    * IF([Version] = "Plan", 1.05 + 0.04 * COS([Line Key] + [MN] * 0.3), 1),
                1
            )
    )'''

VALUE_OR_RATIO = 'IF(SELECTEDVALUE(\'PL Lines\'[Kind]) = "Ratio", "0.0%", "#,##0.0")'
IS_RATIO = "SELECTEDVALUE('PL Lines'[Kind]) = \"Ratio\""

LINE_CALC = r'''
VAR Amt = [PL Amount]
VAR RoomRev = CALCULATE([PL Amount], REMOVEFILTERS('PL Lines'), 'PL Lines'[Line] = "Room Revenue")
VAR FnB = CALCULATE([PL Amount], REMOVEFILTERS('PL Lines'), 'PL Lines'[Line] = "Food & Beverage")
VAR Disc = CALCULATE([PL Amount], REMOVEFILTERS('PL Lines'), 'PL Lines'[Line] = "Discounts & Rebates")
VAR Staff = CALCULATE([PL Amount], REMOVEFILTERS('PL Lines'), 'PL Lines'[Line] = "Staff Cost")
VAR Cogs = CALCULATE([PL Amount], REMOVEFILTERS('PL Lines'), 'PL Lines'[Line] = "Cost of F&B Sold")
VAR Ovh = CALCULATE([PL Amount], REMOVEFILTERS('PL Lines'), 'PL Lines'[Line] = "Overheads")
VAR Rev = RoomRev + FnB + Disc
VAR Contrib = Rev + Staff + Cogs
VAR OpProfit = Contrib + Ovh
RETURN
    SWITCH(
        SELECTEDVALUE('PL Lines'[Line]),
        "Total Revenue", Rev,
        "Contribution", Contrib,
        "Operating Profit", OpProfit,
        "Contribution Margin %", DIVIDE(Contrib, Rev),
        "Staff Cost % of Revenue", DIVIDE(-Staff, Rev),
        "Operating Margin %", DIVIDE(OpProfit, Rev),
        BLANK(), BLANK(),
        Amt
    )'''

LAST_YEAR = r'''CALCULATE(
    [{m}],
    REMOVEFILTERS('PL Calendar'),
    TREATAS(SELECTCOLUMNS(VALUES('PL Calendar'[Month No]), "n", 'PL Calendar'[Month No] - 12), 'PL Calendar'[Month No])
)'''

YTD = r'''VAR LastNo = MAX('PL Calendar'[Month No])
VAR Yr = MAX('PL Calendar'[Year])
RETURN
    CALCULATE([{m}], REMOVEFILTERS('PL Calendar'), 'PL Calendar'[Year] = Yr, 'PL Calendar'[Month No] <= LastNo)'''


def achievement(a, p):
    return f"IF({IS_RATIO}, [{a}] - [{p}], DIVIDE([{a}], [{p}]))"


def growth(a, ly):
    return f"IF({IS_RATIO}, [{a}] - [{ly}], DIVIDE([{a}] - [{ly}], ABS([{ly}])))"


# (name, expression, static format, dynamic format expression or None, folder, hidden)
MEASURES = [
    ("PL Amount", "SUM('PL Fact'[Amount])", "#,##0.0", None, "P&L\\Helpers", True),
    ("PL Line Value", LINE_CALC, None, VALUE_OR_RATIO, "P&L\\Helpers", True),
    ("Actual", "CALCULATE([PL Line Value], 'PL Fact'[Version] = \"Actual\")", None, VALUE_OR_RATIO, "P&L", False),
    ("Plan", "CALCULATE([PL Line Value], 'PL Fact'[Version] = \"Plan\")", None, VALUE_OR_RATIO, "P&L", False),
    ("Last Year", LAST_YEAR.format(m="Actual"), None, VALUE_OR_RATIO, "P&L", False),
    ("vs Plan %", achievement("Actual", "Plan"), "0.0%", None, "P&L", False),
    ("vs LY %", growth("Actual", "Last Year"), "0.0%", None, "P&L", False),
    ("YTD Actual", YTD.format(m="Actual"), None, VALUE_OR_RATIO, "P&L", False),
    ("YTD Plan", YTD.format(m="Plan"), None, VALUE_OR_RATIO, "P&L", False),
    ("YTD Last Year", YTD.format(m="Last Year"), None, VALUE_OR_RATIO, "P&L", False),
    ("YTD vs Plan %", achievement("YTD Actual", "YTD Plan"), "0.0%", None, "P&L", False),
    ("YTD vs LY %", growth("YTD Actual", "YTD Last Year"), "0.0%", None, "P&L", False),
    ("Full Year Plan", "VAR Yr = MAX('PL Calendar'[Year]) RETURN CALCULATE([Plan], REMOVEFILTERS('PL Calendar'), 'PL Calendar'[Year] = Yr)",
     None, VALUE_OR_RATIO, "P&L", False),
    # Same numbers as text, already formatted by the measure
    ("Actual (text)", f"VAR v = [Actual] RETURN IF(NOT ISBLANK(v), FORMAT(v, IF({IS_RATIO}, \"0.0%\", \"#,##0.0\")))", None, None, "P&L", False),
    # Style measures for the visual's formatting wells
    ("Style Row Fill", "IF(SELECTEDVALUE('PL Lines'[Emphasis]) = TRUE(), \"#D9E1F2\")", None, None, "P&L\\Styles", False),
    ("Style Row Bold", "IF(SELECTEDVALUE('PL Lines'[Emphasis]) = TRUE(), 1, 0)", None, None, "P&L\\Styles", False),
    ("Style Fill None", "BLANK()", None, None, "P&L\\Styles", False),
    ("Style Fill None 2", "BLANK()", None, None, "P&L\\Styles", False),
    ("Style Fill Plan", "\"#FFF2CC\"", None, None, "P&L\\Styles", False),
    ("Style Font Below Plan", f"IF(NOT {IS_RATIO} && NOT ISBLANK([Plan]) && [vs Plan %] < 0.95, \"#C00000\")", None, None, "P&L\\Styles", False),
]

NEW_TABLES = ["PL Fact", "PL Lines", "PL Property", "PL Calendar"]
# Earlier test-model names, removed when present
OLD_TABLES = ["MIS Fact", "MIS Particulars", "MIS Unit", "MIS Calendar"]


def calc_table(model, name, expr):
    t = TOM.Table()
    t.Name = name
    p = TOM.Partition()
    p.Name = name
    src = TOM.CalculatedPartitionSource()
    src.Expression = expr.strip()
    p.Source = src
    t.Partitions.Add(p)
    model.Tables.Add(t)
    return t


def col(table, name):
    c = table.Columns.Find(name)
    if c is None:
        raise RuntimeError(f"column {table.Name}[{name}] not found")
    return c


def main():
    server = Server()
    server.Connect(f"Data Source={PORT}")
    model = server.Databases[0].Model

    remove = NEW_TABLES + OLD_TABLES
    for rel in list(model.Relationships):
        if rel.FromTable.Name in remove or rel.ToTable.Name in remove:
            model.Relationships.Remove(rel)
    for name in remove:
        t = model.Tables.Find(name)
        if t is not None:
            model.Tables.Remove(t)
    model.SaveChanges()

    calc_table(model, "PL Lines", LINES)
    calc_table(model, "PL Property", PROPERTIES)
    calc_table(model, "PL Calendar", CALENDAR)
    calc_table(model, "PL Fact", FACT)
    model.SaveChanges()
    model.RequestRefresh(TOM.RefreshType.Calculate)
    model.SaveChanges()

    lines = model.Tables.Find("PL Lines")
    prop = model.Tables.Find("PL Property")
    cal = model.Tables.Find("PL Calendar")
    fact = model.Tables.Find("PL Fact")

    col(lines, "Line").SortByColumn = col(lines, "Line Order")
    col(lines, "Detail").SortByColumn = col(lines, "Line Key")
    for c in ["Line Order", "Base", "Line Key"]:
        col(lines, c).IsHidden = True
    col(prop, "Weight").IsHidden = True
    col(cal, "Month").SortByColumn = col(cal, "Month No")
    for c in ["Month No", "Year Index"]:
        col(cal, c).IsHidden = True
    for c in ["Month Start", "Property", "Line Key", "Amount"]:
        col(fact, c).IsHidden = True

    h = TOM.Hierarchy()
    h.Name = "P&L"
    for i, c in enumerate(["Line", "Detail"]):
        lvl = TOM.Level()
        lvl.Name = c
        lvl.Ordinal = i
        lvl.Column = col(lines, c)
        h.Levels.Add(lvl)
    lines.Hierarchies.Add(h)

    def rel(from_t, from_c, to_t, to_c):
        r = TOM.SingleColumnRelationship()
        r.Name = f"PL {from_t.Name}.{from_c} to {to_t.Name}.{to_c}"
        r.FromColumn = col(from_t, from_c)
        r.ToColumn = col(to_t, to_c)
        r.FromCardinality = TOM.RelationshipEndCardinality.Many
        r.ToCardinality = TOM.RelationshipEndCardinality.One
        model.Relationships.Add(r)

    rel(fact, "Month Start", cal, "Month Start")
    rel(fact, "Property", prop, "Property")
    rel(fact, "Line Key", lines, "Line Key")

    for name, expr, fmt, dyn, folder, hidden in MEASURES:
        m = TOM.Measure()
        m.Name = name
        m.Expression = expr.strip()
        # A measure may have a static or a dynamic format string, not both
        if dyn:
            fd = TOM.FormatStringDefinition()
            fd.Expression = dyn
            m.FormatStringDefinition = fd
        elif fmt:
            m.FormatString = fmt
        m.DisplayFolder = folder
        m.IsHidden = hidden
        fact.Measures.Add(m)

    model.SaveChanges()
    model.RequestRefresh(TOM.RefreshType.Calculate)
    model.SaveChanges()
    print("P&L test model created:", ", ".join(NEW_TABLES), "|", len(MEASURES), "measures")
    server.Disconnect()


if __name__ == "__main__":
    main()
