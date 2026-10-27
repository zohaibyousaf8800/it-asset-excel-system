#!/usr/bin/env python3
"""
IT Asset & Helpdesk Management System — workbook generator.

Builds a fully formatted, formula-driven Excel workbook for IT operations:
asset inventory, helpdesk tickets, software licence tracking and new-starter
onboarding — with a live dashboard driven entirely by Excel formulas.

Author: Zohaib Yousaf (info@zohaibyousaf.com)
Requires: openpyxl  (pip install openpyxl)
Run:      python build_workbook.py
Output:   IT_Asset_Helpdesk_Management_System.xlsx
"""
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT = Path(__file__).with_name("IT_Asset_Helpdesk_Management_System.xlsx")

# ---------------------------------------------------------------- theme
NAVY = "0F2A44"
TEAL = "0E7C7B"
LIGHT = "E8EEF3"
WHITE = "FFFFFF"
GREY_TEXT = "475569"
GOLD = "F2B134"

header_font = Font(name="Calibri", size=11, bold=True, color=WHITE)
header_fill = PatternFill("solid", fgColor=NAVY)
title_font = Font(name="Calibri", size=18, bold=True, color=NAVY)
sub_font = Font(name="Calibri", size=11, color=GREY_TEXT)
kpi_label_font = Font(name="Calibri", size=11, bold=True, color=GREY_TEXT)
kpi_value_font = Font(name="Calibri", size=22, bold=True, color=NAVY)
thin = Side(style="thin", color="CBD5E1")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(vertical="center", wrap_text=True)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
    ws.row_dimensions[row].height = 30


def add_table(ws, ref, name):
    tab = Table(displayName=name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True, showFirstColumn=False
    )
    ws.add_table(tab)


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ---------------------------------------------------------------- data
TODAY = date(2026, 10, 27)

ASSETS = [
    # id, type, model, serial, user, dept, location, bought, warranty, status, compliant
    ("AST-001", "Laptop", "Dell Latitude 5540", "DL5540-8841", "A. Khan", "Finance", "Islamabad HQ", "2024-03-12", "2027-03-12", "In Use", "Yes"),
    ("AST-002", "Laptop", "HP EliteBook 840 G10", "HP840-2210", "S. Ahmed", "HR", "Islamabad HQ", "2024-06-01", "2027-06-01", "In Use", "Yes"),
    ("AST-003", "Laptop", "Lenovo ThinkPad T14", "LT14-9034", "Unassigned", "IT", "IT Store", "2023-11-20", "2026-11-20", "In Stock", "Yes"),
    ("AST-004", "Desktop", "Dell OptiPlex 7010", "DO7010-117", "M. Raza", "Operations", "Islamabad HQ", "2023-02-14", "2026-02-14", "In Use", "Yes"),
    ("AST-005", "Monitor", "Dell P2422H 24in", "DP24-5521", "A. Khan", "Finance", "Islamabad HQ", "2024-03-12", "2027-03-12", "In Use", "Yes"),
    ("AST-006", "Laptop", "Dell Latitude 5540", "DL5540-8850", "F. Noor", "Marketing", "Remote", "2024-09-05", "2027-09-05", "In Use", "No"),
    ("AST-007", "Mobile", "Samsung Galaxy A54", "SGA54-331", "Field Team 1", "Operations", "Field", "2024-01-18", "2026-01-18", "In Use", "Yes"),
    ("AST-008", "Printer", "HP LaserJet Pro M404", "HPLJ-078", "Shared — Floor 2", "Admin", "Islamabad HQ", "2022-08-30", "2025-08-30", "Under Repair", "Yes"),
    ("AST-009", "Laptop", "HP ProBook 450 G9", "HP450-664", "R. Iqbal", "Finance", "Islamabad HQ", "2023-05-22", "2026-05-22", "In Use", "Yes"),
    ("AST-010", "Server", "Dell PowerEdge R750", "DPR750-01", "Data Centre", "IT", "Server Room", "2023-07-01", "2028-07-01", "In Use", "Yes"),
    ("AST-011", "Network Switch", "Cisco Catalyst 9200", "CC9200-14", "Network Rack A", "IT", "Server Room", "2023-07-01", "2028-07-01", "In Use", "Yes"),
    ("AST-012", "Laptop", "Lenovo IdeaPad 3", "LI3-209", "Unassigned", "IT", "IT Store", "2021-04-10", "2024-04-10", "Retired", "No"),
    ("AST-013", "Tablet", "Apple iPad 10th Gen", "IP10-771", "Sales Team", "Sales", "Islamabad HQ", "2025-02-11", "2027-02-11", "In Use", "Yes"),
    ("AST-014", "Laptop", "Dell Latitude 7440", "DL7440-19", "Z. Yousaf", "IT", "Islamabad HQ", "2025-01-15", "2028-01-15", "In Use", "Yes"),
    ("AST-015", "Desktop", "HP ProDesk 400 G9", "HP400-318", "H. Tariq", "HR", "Islamabad HQ", "2024-10-02", "2027-10-02", "In Use", "Yes"),
    ("AST-016", "Monitor", "Samsung 27in Curved", "SM27-940", "Design Desk", "Marketing", "Islamabad HQ", "2024-04-19", "2027-04-19", "In Use", "Yes"),
    ("AST-017", "Laptop", "Microsoft Surface Pro 9", "MSP9-055", "Management", "Management", "Islamabad HQ", "2024-12-01", "2027-12-01", "In Use", "Yes"),
    ("AST-018", "Mobile", "Apple iPhone 13", "IP13-612", "Management", "Management", "Remote", "2023-09-09", "2025-09-09", "In Use", "Yes"),
    ("AST-019", "Laptop", "HP EliteBook 840 G9", "HP840-198", "Unassigned", "IT", "IT Store", "2023-08-25", "2026-08-25", "Under Repair", "No"),
    ("AST-020", "UPS", "APC Smart-UPS 1500", "APC1500-07", "Server Room", "IT", "Server Room", "2024-05-30", "2027-05-30", "In Use", "Yes"),
]

TICKETS = [
    # id, logged, requester, dept, category, priority, description, tech, status, resolved, hours
    ("TCK-1001", "2026-10-20", "A. Khan", "Finance", "Microsoft 365", "High", "Shared mailbox permissions not applying in Outlook", "Z. Yousaf", "Resolved", "2026-10-20", 3.5),
    ("TCK-1002", "2026-10-21", "S. Ahmed", "HR", "Hardware", "Medium", "Laptop battery swelling — replacement required", "Z. Yousaf", "In Progress", None, None),
    ("TCK-1003", "2026-10-21", "M. Raza", "Operations", "Network", "Critical", "VPN drops every 10 minutes for remote users", "Z. Yousaf", "Resolved", "2026-10-22", 9.0),
    ("TCK-1004", "2026-10-22", "F. Noor", "Marketing", "Microsoft 365", "Medium", "OneDrive sync conflicts on shared folder", "Support Team", "Resolved", "2026-10-22", 2.0),
    ("TCK-1005", "2026-10-22", "R. Iqbal", "Finance", "Security", "High", "MFA prompt loop after password change", "Z. Yousaf", "Resolved", "2026-10-23", 4.5),
    ("TCK-1006", "2026-10-23", "H. Tariq", "HR", "Access", "Medium", "New starter account and Intune enrolment", "Support Team", "Closed", "2026-10-24", 6.0),
    ("TCK-1007", "2026-10-23", "Field Team 1", "Operations", "Hardware", "Low", "Phone case and charger replacement", "Support Team", "Closed", "2026-10-24", 1.0),
    ("TCK-1008", "2026-10-24", "Sales Team", "Sales", "Software", "Medium", "CRM client crashing on launch", "Z. Yousaf", "In Progress", None, None),
    ("TCK-1009", "2026-10-24", "Design Desk", "Marketing", "Hardware", "Low", "Second monitor flickering at 60Hz", "Support Team", "Open", None, None),
    ("TCK-1010", "2026-10-25", "Management", "Management", "Security", "Critical", "Conditional Access policy review after failed sign-ins", "Z. Yousaf", "In Progress", None, None),
    ("TCK-1011", "2026-10-25", "A. Khan", "Finance", "Software", "Medium", "Excel reports slow on large workbook", "Support Team", "Resolved", "2026-10-26", 3.0),
    ("TCK-1012", "2026-10-26", "Shared — Floor 2", "Admin", "Hardware", "High", "Printer offline — print spooler fault", "Support Team", "Open", None, None),
    ("TCK-1013", "2026-10-26", "S. Ahmed", "HR", "Microsoft 365", "Low", "Teams meeting room calendar not visible", "Support Team", "Open", None, None),
    ("TCK-1014", "2026-10-26", "M. Raza", "Operations", "Network", "Medium", "Wi-Fi dead spot in meeting room 2", "Z. Yousaf", "In Progress", None, None),
    ("TCK-1015", "2026-10-27", "New Starter", "HR", "Access", "Medium", "Onboarding — device, M365 licence and groups", "Z. Yousaf", "Open", None, None),
]

LICENSES = [
    # software, vendor, total, assigned, renewal, cost each
    ("Microsoft 365 Business Premium", "Microsoft", 85, 78, "2027-03-01", 22.00),
    ("Microsoft 365 Business Standard", "Microsoft", 40, 34, "2027-03-01", 12.50),
    ("Windows Server 2022 (cores)", "Microsoft", 16, 16, "2027-07-01", 90.00),
    ("Adobe Acrobat Pro", "Adobe", 12, 10, "2026-12-15", 14.99),
    ("Antivirus — Endpoint Protection", "Vendor", 120, 112, "2027-01-31", 3.50),
    ("Backup & Recovery Agents", "Vendor", 30, 27, "2027-05-30", 8.00),
]

ONBOARDING = [
    ("Create Microsoft 365 account & assign licence", "IT", "Done"),
    ("Add user to security groups & shared mailboxes", "IT", "Done"),
    ("Enrol device in Intune & apply compliance policy", "IT", "Done"),
    ("Configure MFA & Conditional Access", "IT", "In Progress"),
    ("Install standard apps (Office, Teams, VPN)", "IT", "In Progress"),
    ("Hand over device & asset sign-off (Assets sheet)", "IT / HR", "Not Started"),
    ("Security awareness briefing & password policy", "IT", "Not Started"),
    ("Add to HR system & payroll", "HR", "Done"),
    ("Manager intro & first-week plan", "HR / Manager", "Not Started"),
]


def d(s):
    return date.fromisoformat(s) if s else None


# ---------------------------------------------------------------- build
wb = Workbook()

# ---- Assets
ws = wb.active
ws.title = "Assets"
ws.sheet_properties.tabColor = NAVY
headers = ["Asset ID", "Device Type", "Brand / Model", "Serial Number", "Assigned To",
           "Department", "Location", "Purchase Date", "Warranty Expiry",
           "Warranty (days left)", "Status", "Intune Compliant"]
ws.append(headers)
for a in ASSETS:
    ws.append([a[0], a[1], a[2], a[3], a[4], a[5], a[6], d(a[7]), d(a[8]),
               f"=I{ws.max_row + 1}-TODAY()", a[9], a[10]])
style_header(ws, 1, len(headers))
last = ws.max_row
for r in range(2, last + 1):
    for c in (8, 9):
        ws.cell(row=r, column=c).number_format = "yyyy-mm-dd"
    ws.cell(row=r, column=10).alignment = center
    ws.cell(row=r, column=11).alignment = center
    ws.cell(row=r, column=12).alignment = center
add_table(ws, f"A1:L{last}", "AssetsTable")
set_widths(ws, [11, 15, 26, 15, 17, 13, 14, 13, 14, 15, 13, 14])
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:L{last}"
ws.conditional_formatting.add(f"J2:J{last}",
    CellIsRule(operator="between", formula=["0", "90"], fill=PatternFill("solid", fgColor="FDE68A")))
ws.conditional_formatting.add(f"J2:J{last}",
    CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="FECACA")))
dv = DataValidation(type="list", formula1='"In Use,In Stock,Under Repair,Retired"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f"K2:K500")
dv2 = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
ws.add_data_validation(dv2); dv2.add(f"L2:L500")

# ---- Tickets
ws = wb.create_sheet("Helpdesk_Tickets")
ws.sheet_properties.tabColor = TEAL
headers = ["Ticket ID", "Date Logged", "Requester", "Department", "Category", "Priority",
           "Description", "Assigned Technician", "Status", "Resolution Date",
           "Resolution (hrs)", "SLA Breach (48h)"]
ws.append(headers)
for t in TICKETS:
    r = ws.max_row + 1
    ws.append([t[0], d(t[1]), t[2], t[3], t[4], t[5], t[6], t[7], t[8],
               d(t[9]) if t[9] else None, t[10],
               f'=IF(AND(K{r}<>"",K{r}>48),"YES","")'])
style_header(ws, 1, len(headers))
last = ws.max_row
for r in range(2, last + 1):
    ws.cell(row=r, column=2).number_format = "yyyy-mm-dd"
    ws.cell(row=r, column=10).number_format = "yyyy-mm-dd"
    for c in (1, 6, 9, 11, 12):
        ws.cell(row=r, column=c).alignment = center
add_table(ws, f"A1:L{last}", "TicketsTable")
set_widths(ws, [11, 12, 16, 13, 15, 10, 46, 17, 12, 14, 14, 14])
ws.freeze_panes = "A2"
ws.conditional_formatting.add(f"F2:F{last}",
    FormulaRule(formula=['$F2="Critical"'], fill=PatternFill("solid", fgColor="FECACA")))
ws.conditional_formatting.add(f"F2:F{last}",
    FormulaRule(formula=['$F2="High"'], fill=PatternFill("solid", fgColor="FDE68A")))
dv = DataValidation(type="list", formula1='"Open,In Progress,Resolved,Closed"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f"I2:I500")
dv = DataValidation(type="list", formula1='"Low,Medium,High,Critical"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f"F2:F500")

# ---- Licences
ws = wb.create_sheet("Licenses")
ws.sheet_properties.tabColor = GOLD
headers = ["Software", "Vendor", "Licences Owned", "Assigned", "Available",
           "Utilisation %", "Renewal Date", "Cost / licence (USD/mo)", "Monthly Cost (USD)", "Status"]
ws.append(headers)
for lic in LICENSES:
    r = ws.max_row + 1
    ws.append([lic[0], lic[1], lic[2], lic[3], f"=C{r}-D{r}", f"=IF(C{r}=0,0,D{r}/C{r})",
               d(lic[4]), lic[5], f"=D{r}*H{r}",
               f'=IF(E{r}<=0,"FULL — renew / add",IF(E{r}<=5,"Low — plan renewal","OK"))'])
style_header(ws, 1, len(headers))
last = ws.max_row
for r in range(2, last + 1):
    ws.cell(row=r, column=6).number_format = "0%"
    ws.cell(row=r, column=7).number_format = "yyyy-mm-dd"
    for c in (8, 9):
        ws.cell(row=r, column=c).number_format = '"$"#,##0.00'
add_table(ws, f"A1:J{last}", "LicensesTable")
set_widths(ws, [36, 12, 14, 10, 10, 13, 13, 18, 16, 22])
ws.freeze_panes = "A2"

# ---- Onboarding
ws = wb.create_sheet("Onboarding_Checklist")
ws.sheet_properties.tabColor = "64748B"
ws.append(["New Starter Onboarding — IT Checklist", "", "", ""])
ws["A1"].font = title_font
ws.append(["Task", "Owner", "Status", "Notes"])
style_header(ws, 2, 4)
for task, owner, status in ONBOARDING:
    ws.append([task, owner, status, ""])
add_table(ws, f"A2:D{ws.max_row}", "OnboardingTable")
set_widths(ws, [52, 14, 14, 30])
dv = DataValidation(type="list", formula1='"Not Started,In Progress,Done"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f"C3:C100")

# ---- Summary helpers for charts (hidden-ish small block on Dashboard data sheet)
ws = wb.create_sheet("Summary")
ws.append(["Ticket Status", "Count"])
for s in ["Open", "In Progress", "Resolved", "Closed"]:
    ws.append([s, f"=COUNTIF(Helpdesk_Tickets!I:I,A{ws.max_row + 1})"])
ws.append([])
ws.append(["Asset Status", "Count"])
for s in ["In Use", "In Stock", "Under Repair", "Retired"]:
    ws.append([s, f"=COUNTIF(Assets!K:K,A{ws.max_row + 1})"])
set_widths(ws, [16, 8])

# ---- Dashboard (first sheet)
ws = wb.create_sheet("Dashboard", 0)
ws.sheet_properties.tabColor = TEAL
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws["B2"] = "IT Asset & Helpdesk Management System"
ws["B2"].font = title_font
ws["B3"] = "Built by Zohaib Yousaf — IT Manager  |  info@zohaibyousaf.com  |  All figures below are live Excel formulas."
ws["B3"].font = sub_font
ws.merge_cells("B3:I3")

kpis = [
    ("Total Assets", "=COUNTA(Assets!A2:A500)"),
    ("Open / Active Tickets", '=COUNTIF(Helpdesk_Tickets!I:I,"Open")+COUNTIF(Helpdesk_Tickets!I:I,"In Progress")'),
    ("Device Compliance %", '=TEXT(COUNTIFS(Assets!L:L,"Yes",Assets!K:K,"<>Retired")/MAX(1,COUNTA(Assets!L2:L500)-COUNTIF(Assets!K:K,"Retired")),"0%")'),
    ("Avg Resolution (hrs)", "=ROUND(AVERAGE(Helpdesk_Tickets!K:K),1)"),
    ("Licences Available", "=SUM(Licenses!E:E)"),
    ("Monthly Licence Cost", '=TEXT(SUM(Licenses!I:I),"$#,##0")'),
]
row = 5
for i, (label, formula) in enumerate(kpis):
    col = 2 + (i % 3) * 3
    r = row + (i // 3) * 4
    lbl = ws.cell(row=r, column=col, value=label); lbl.font = kpi_label_font
    val = ws.cell(row=r + 1, column=col, value=formula); val.font = kpi_value_font
    for rr in (r, r + 1):
        for cc in range(col, col + 2):
            ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor=LIGHT)
            ws.cell(row=rr, column=cc).border = border

ws["B14"] = "How to use:  1) Keep Assets, Helpdesk_Tickets and Licenses updated — the Dashboard recalculates itself.  " \
            "2) Warranty column flags cover expiring in 90 days (amber) and expired (red).  " \
            "3) Use the dropdowns — do not type statuses freehand.  4) Onboarding_Checklist tracks every new starter."
ws["B14"].font = sub_font; ws["B14"].alignment = wrap
ws.merge_cells("B14:I17")
set_widths(ws, [2, 18, 18, 4, 18, 18, 4, 18, 18])

bar = BarChart(); bar.type = "col"; bar.style = 10
bar.title = "Tickets by Status"; bar.y_axis.title = "Tickets"; bar.x_axis.title = "Status"
data = Reference(wb["Summary"], min_col=2, min_row=1, max_row=5)
cats = Reference(wb["Summary"], min_col=1, min_row=2, max_row=5)
bar.add_data(data, titles_from_data=True); bar.set_categories(cats)
bar.height, bar.width = 7.5, 13
ws.add_chart(bar, "B18")

pie = PieChart(); pie.style = 10; pie.title = "Assets by Status"
data = Reference(wb["Summary"], min_col=2, min_row=7, max_row=11)
cats = Reference(wb["Summary"], min_col=1, min_row=8, max_row=11)
pie.add_data(data, titles_from_data=True); pie.set_categories(cats)
pie.height, pie.width = 7.5, 13
ws.add_chart(pie, "F18")

for name in wb.sheetnames:
    wb[name].sheet_properties.pageSetUpPr.fitToPage = True

wb.properties.creator = "Zohaib Yousaf"
wb.properties.title = "IT Asset & Helpdesk Management System"
wb.properties.description = "Formula-driven IT asset, helpdesk and licence management workbook by Zohaib Yousaf, IT Manager."
wb.save(OUT)
print(f"Saved {OUT} ({OUT.stat().st_size} bytes)")
