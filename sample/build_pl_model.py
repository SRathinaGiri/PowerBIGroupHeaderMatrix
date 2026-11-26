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

# Chart of accounts: Line -> Detail -> GL accounts (code, name, monthly base amount in thousands).
# Lines without details are calculated (subtotals and ratios); they get a single Detail/GL row.
# (line, kind, emphasis, [(detail, [(gl_code, gl_name, base), ...]), ...])
CHART = [
    ("Room Revenue", "Value", False, [
        ("Leisure Stays", [(4101, "Direct Bookings", 260), (4102, "Online Travel Agents", 340), (4103, "Package Holidays", 220)]),
        ("Corporate Stays", [(4111, "Corporate Contracts", 300), (4112, "Government & Embassy", 90), (4113, "Airline Crew", 150)]),
    ]),
    ("Food & Beverage", "Value", False, [
        ("Restaurants", [(4201, "All-Day Dining", 140), (4202, "Speciality Restaurant", 95), (4203, "In-Room Dining", 75)]),
        ("Banquets & Events", [(4211, "Weddings", 120), (4212, "Conferences", 95), (4213, "Outdoor Catering", 45)]),
    ]),
    ("Discounts & Rebates", "Value", False, [
        ("Loyalty Rebates", [(4301, "Points Redemption", -25), (4302, "Member Discounts", -20)]),
        ("Promotional Offers", [(4311, "Seasonal Offers", -22), (4312, "Corporate Rebates", -16)]),
    ]),
    ("Total Revenue", "Value", True, []),
    ("Staff Cost", "Value", False, [
        ("Front Office", [(5101, "Front Office Wages", -140), (5102, "Front Office Overtime", -28), (5103, "Front Office Training", -12), (5104, "Uniforms", -10)]),
        ("Housekeeping", [(5111, "Housekeeping Wages", -115), (5112, "Contract Labour", -35), (5113, "Housekeeping Supplies", -10)]),
        ("Kitchen", [(5121, "Chefs & Cooks", -85), (5122, "Stewarding", -25), (5123, "Staff Meals", -10)]),
    ]),
    ("Cost of F&B Sold", "Value", False, [
        ("Food Cost", [(5201, "Meat & Seafood", -60), (5202, "Dairy & Produce", -45), (5203, "Dry Goods", -45)]),
        ("Beverage Cost", [(5211, "Wines & Spirits", -32), (5212, "Soft Drinks", -14), (5213, "Coffee & Tea", -9)]),
    ]),
    ("Contribution", "Value", True, []),
    ("Contribution Margin %", "Ratio", False, []),
    ("Staff Cost % of Revenue", "Ratio", False, []),
    ("Overheads", "Value", False, [
        ("Energy", [(6101, "Electricity", -55), (6102, "Water", -15), (6103, "Gas & Fuel", -25)]),
        ("Maintenance", [(6111, "Repairs", -30), (6112, "Service Contracts", -22), (6113, "Landscaping", -18)]),
        ("Lease & Admin", [(6121, "Lease Rent", -70), (6122, "Insurance", -20), (6123, "IT & Software", -18), (6124, "Professional Fees", -22)]),
    ]),
    ("Operating Profit", "Value", True, []),
    ("Operating Margin %", "Ratio", False, []),
]


def build_lines_table() -> str:
    rows = []
    detail_order = 0
    for line_order, (line, kind, emphasis, details) in enumerate(CHART, start=1):
        flag = "TRUE" if emphasis else "FALSE"
        if not details:
            detail_order += 1
            key = 9000 + line_order  # calculated line: one Detail and one GL row named after it
            rows.append(f'{{ {key}, "{line}", {detail_order}, "{line}", "{line}", {line_order}, "{kind}", {flag}, 0 }}')
            continue
        for detail, accounts in details:
            detail_order += 1
            for code, name, base in accounts:
                rows.append(f'{{ {code}, "{code} - {name}", {detail_order}, "{detail}", "{line}", {line_order}, "{kind}", {flag}, {base} }}')
    body = ",\n        ".join(rows)
    return f"""
DATATABLE(
    "GL Key", INTEGER, "GL Code", STRING, "Detail Order", INTEGER, "Detail", STRING, "Line", STRING,
    "Line Order", INTEGER, "Kind", STRING, "Emphasis", BOOLEAN, "Base", DOUBLE,
    {{
        {body}
    }}
)"""


LINES = build_lines_table()

# Weights add up to 1 within a region; Scale makes the regions differ in size
PROPERTIES = r'''
DATATABLE(
    "Property No", INTEGER, "Property", STRING, "Region", STRING, "City", STRING, "Segment", STRING, "Weight", DOUBLE, "Scale", DOUBLE,
    {
        { 1, "Harbour View Hotel", "North", "Lakeport", "Urban", 0.34, 1.00 },
        { 2, "Lakeport Airport Inn", "North", "Lakeport", "Urban", 0.22, 1.00 },
        { 3, "Pine Ridge Lodge", "North", "Pine Ridge", "Leisure", 0.18, 1.00 },
        { 4, "Northgate Suites", "North", "Northgate", "Urban", 0.26, 1.00 },
        { 5, "Central Plaza Hotel", "Central", "Midvale", "Urban", 0.31, 0.80 },
        { 6, "Midvale Business Hotel", "Central", "Midvale", "Urban", 0.24, 0.80 },
        { 7, "Riverside Retreat", "Central", "Riverbend", "Leisure", 0.20, 0.80 },
        { 8, "Old Town Boutique", "Central", "Riverbend", "Urban", 0.25, 0.80 },
        { 9, "Sunset Beach Resort", "South", "Coral Bay", "Leisure", 0.36, 1.25 },
        { 10, "Coral Bay Villas", "South", "Coral Bay", "Leisure", 0.22, 1.25 },
        { 11, "Southport Conference Hotel", "South", "Southport", "Urban", 0.27, 1.25 },
        { 12, "Palm Grove Inn", "South", "Southport", "Leisure", 0.15, 1.25 }
    }
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

# Each property gets its own seasonal phase and cost ratio; leisure properties swing more.
# Plan differs from Actual by roughly +/-7% depending on account group, month and property.
FACT = r'''
VAR Months = SELECTCOLUMNS('PL Calendar', "Month Start", 'PL Calendar'[Month Start], "MN", 'PL Calendar'[Month No], "YI", 'PL Calendar'[Year Index])
VAR Accounts = SELECTCOLUMNS(FILTER('PL Lines', 'PL Lines'[Base] <> 0), "GL Key", 'PL Lines'[GL Key], "B", 'PL Lines'[Base])
VAR Props = SELECTCOLUMNS('PL Property', "Property", 'PL Property'[Property], "PN", 'PL Property'[Property No],
    "W", 'PL Property'[Weight] * 'PL Property'[Scale], "Swing", IF('PL Property'[Segment] = "Leisure", 0.22, 0.10))
VAR Versions = DATATABLE("Version", STRING, { { "Actual" }, { "Plan" } })
VAR AllRows = GENERATE(GENERATE(GENERATE(Months, Accounts), Props), Versions)
VAR Kept = FILTER(AllRows, [Version] = "Plan" || [Month Start] <= DATE(2026, 8, 1))
RETURN
    SELECTCOLUMNS(
        Kept,
        "Month Start", [Month Start],
        "Property", [Property],
        "GL Key", [GL Key],
        "Version", [Version],
        "Amount",
            ROUND(
                [B] * [W]
                    * (1 + [Swing] * SIN([MN] * 0.52 + [PN] * 1.3))
                    * (1 + 0.05 * SIN([GL Key] * 0.37 + [MN] * 0.9))
                    * POWER(1.07 + [PN] * 0.004, [YI])
                    * IF([GL Key] >= 5000 && [GL Key] < 9000, 0.86 + MOD([PN] * 7, 10) * 0.03, 1)
                    * IF([Version] = "Plan",
                        1 + 0.06 * SIN([MN] * 0.6 + INT([GL Key] / 100) * 0.8) + 0.03 * SIN([PN] * 1.9 + [MN]),
                        1),
                2
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


# Green when [actual] beats [target], red when it misses.
# Amounts are signed (costs negative), so actual >= target is better for revenue, profit and costs alike.
# Ratios: higher margins are better; a higher staff-cost share is worse.
FONT_BETTER_WORSE = r'''
VAR A = [{actual}]
VAR T = [{target}]
VAR Better =
    SWITCH(
        TRUE(),
        SELECTEDVALUE('PL Lines'[Line]) = "Staff Cost % of Revenue", A <= T,
        A >= T
    )
RETURN
    IF(NOT ISBLANK(A) && NOT ISBLANK(T), IF(Better, "#1E7B34", "#C00000"))'''


def font_rule(actual: str, target: str) -> str:
    return FONT_BETTER_WORSE.replace("{actual}", actual).replace("{target}", target)


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
    ("Style Fill Plan", "\"#FFF2CC\"", None, None, "P&L\\Styles", False),
    # Green/red by comparison; map each to its column in Format > Measure columns > Cell font color from
    ("Style Font vs Plan", font_rule("Actual", "Plan"), None, None, "P&L\\Styles", False),
    ("Style Font vs LY", font_rule("Actual", "Last Year"), None, None, "P&L\\Styles", False),
    ("Style Font YTD vs Plan", font_rule("YTD Actual", "YTD Plan"), None, None, "P&L\\Styles", False),
    ("Style Font YTD vs LY", font_rule("YTD Actual", "YTD Last Year"), None, None, "P&L\\Styles", False),
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
    col(lines, "Detail").SortByColumn = col(lines, "Detail Order")
    col(lines, "GL Code").SortByColumn = col(lines, "GL Key")
    for c in ["Line Order", "Detail Order", "Base", "GL Key"]:
        col(lines, c).IsHidden = True
    for c in ["Weight", "Scale", "Property No"]:
        col(prop, c).IsHidden = True
    col(cal, "Month").SortByColumn = col(cal, "Month No")
    for c in ["Month No", "Year Index"]:
        col(cal, c).IsHidden = True
    for c in ["Month Start", "Property", "GL Key", "Amount"]:
        col(fact, c).IsHidden = True

    def hierarchy(table, name, columns):
        h = TOM.Hierarchy()
        h.Name = name
        for i, c in enumerate(columns):
            lvl = TOM.Level()
            lvl.Name = c
            lvl.Ordinal = i
            lvl.Column = col(table, c)
            h.Levels.Add(lvl)
        table.Hierarchies.Add(h)

    hierarchy(lines, "P&L", ["Line", "Detail", "GL Code"])
    hierarchy(prop, "Location", ["Region", "City", "Property"])
    hierarchy(cal, "Period", ["Year", "Quarter", "Month"])

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
    rel(fact, "GL Key", lines, "GL Key")

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
