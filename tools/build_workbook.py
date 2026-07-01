"""
Generates the Arabic small-business management workbook (Business_Workbook.xlsx).

Sheets:
  - لوحة التحكم  (Dashboard, KPIs + charts)
  - المبيعات      (Sales table + weekly/monthly/grand totals)
  - المصروفات     (Expenses table + category/monthly totals)
  - الموظفون      (Payroll summary [dynamic array] + Work Log table + Transactions table)
  - الإعدادات     (Hidden master-data sheet: employees, categories, payment methods, company settings)

Run: python3 tools/build_workbook.py
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.styles.protection import Protection
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.properties import CalcProperties
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.data_source import AxDataSource, StrRef
from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties

# ----------------------------------------------------------------------------
# Theme
# ----------------------------------------------------------------------------
NAVY = "1F3B57"
TEAL = "12766A"
GOLD = "C99A2E"
RED = "B0392A"
LIGHT_BG = "F5F7FA"
BAND = "EEF2F5"
EDIT_FILL = "FFF6DB"      # soft yellow -> editable cells
LOCK_FILL = "F1F3F5"      # soft gray  -> computed / locked cells
WHITE = "FFFFFF"
GRID = "D6DCE1"

FONT_NAME = "Calibri"

f_title      = Font(name=FONT_NAME, size=18, bold=True, color=WHITE)
f_subtitle   = Font(name=FONT_NAME, size=10, italic=True, color="6B7684")
f_section    = Font(name=FONT_NAME, size=13, bold=True, color=WHITE)
f_header     = Font(name=FONT_NAME, size=11, bold=True, color=WHITE)
f_normal     = Font(name=FONT_NAME, size=11, color="1B1B1B")
f_kpi_label  = Font(name=FONT_NAME, size=11, bold=True, color=WHITE)
f_kpi_value  = Font(name=FONT_NAME, size=16, bold=True, color=WHITE)
f_note       = Font(name=FONT_NAME, size=9, italic=True, color="8A93A0")

fill_title   = PatternFill("solid", fgColor=NAVY)
fill_section = PatternFill("solid", fgColor=TEAL)
fill_header  = PatternFill("solid", fgColor=NAVY)
fill_band    = PatternFill("solid", fgColor=BAND)
fill_edit    = PatternFill("solid", fgColor=EDIT_FILL)
fill_lock    = PatternFill("solid", fgColor=LOCK_FILL)
fill_page    = PatternFill("solid", fgColor=LIGHT_BG)

thin = Side(style="thin", color=GRID)
border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

LOCKED = Protection(locked=True)
UNLOCKED = Protection(locked=False)

CUR_FMT = '#,##0.00" ر.س"'
DATE_FMT = "yyyy/mm/dd"
TIME_FMT = "hh:mm AM/PM"
HOURS_FMT = '0.00" ساعة"'

DASH, SALES, EXP, EMP, SET = (
    "لوحة التحكم", "المبيعات", "المصروفات", "الموظفون", "الإعدادات",
)


def set_col_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def rtl(ws):
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False


def title_bar(ws, row, col_start, col_end, text, fill=fill_title, font=f_title, height=30):
    ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
    c = ws.cell(row=row, column=col_start, value=text)
    c.font = font
    c.fill = fill
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = height
    for col in range(col_start, col_end + 1):
        ws.cell(row=row, column=col).fill = fill
        ws.cell(row=row, column=col).protection = LOCKED


def header_cell(ws, row, col, text):
    c = ws.cell(row=row, column=col, value=text)
    c.font = f_header
    c.fill = fill_header
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = border_all
    c.protection = LOCKED
    return c


def note(ws, row, col_start, col_end, text):
    ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
    c = ws.cell(row=row, column=col_start, value=text)
    c.font = f_note
    c.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[row].height = 14


def style_cell(ws, row, col, editable=False, number_format=None, align="center"):
    c = ws.cell(row=row, column=col)
    c.font = f_normal
    c.border = border_all
    c.alignment = Alignment(horizontal=align, vertical="center")
    if number_format:
        c.number_format = number_format
    if editable:
        c.fill = fill_edit
        c.protection = UNLOCKED
    else:
        c.fill = fill_lock
        c.protection = LOCKED
    return c


def add_defined_name(wb, name, formula):
    wb.defined_names[name] = DefinedName(name, attr_text=formula)


def add_table(ws, name, ref, style="TableStyleLight9", bands=False):
    tbl = Table(displayName=name, ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(
        name=style, showFirstColumn=False, showLastColumn=False,
        showRowStripes=bands, showColumnStripes=False,
    )
    ws.add_table(tbl)
    return tbl


def list_dv(formula1, prompt_title, prompt_msg):
    dv = DataValidation(type="list", formula1=formula1, allow_blank=True, showDropDown=False)
    dv.promptTitle = prompt_title
    dv.prompt = prompt_msg
    dv.showInputMessage = True
    dv.showErrorMessage = True
    dv.errorTitle = "قيمة غير صحيحة"
    dv.error = "الرجاء الاختيار من القائمة المتاحة فقط."
    return dv


def set_text_categories(chart, ref):
    """openpyxl's set_categories() always emits numRef; text-valued category
    axes (labels, category names) need strRef or Excel/LO show blank labels."""
    for s in chart.series:
        s.cat = AxDataSource(strRef=StrRef(f=str(ref)))


def positive_number_dv(prompt_title):
    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                         allow_blank=True, showErrorMessage=True)
    dv.promptTitle = prompt_title
    dv.prompt = "أدخل رقمًا أكبر من أو يساوي صفر."
    dv.errorTitle = "قيمة غير صحيحة"
    dv.error = "الرجاء إدخال رقم صحيح أكبر من أو يساوي صفر."
    return dv


# ============================================================================
wb = Workbook()
wb.calculation = CalcProperties(fullCalcOnLoad=True)
wb.properties.title = "دفتر إدارة الأعمال"
wb.properties.creator = "Business Workbook Generator"

ws_dash = wb.active
ws_dash.title = DASH
ws_sales = wb.create_sheet(SALES)
ws_exp = wb.create_sheet(EXP)
ws_emp = wb.create_sheet(EMP)
ws_set = wb.create_sheet(SET)

for ws in (ws_dash, ws_sales, ws_exp, ws_emp, ws_set):
    rtl(ws)
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# ============================================================================
# SETTINGS SHEET (hidden master data)
# ============================================================================
ws = ws_set
set_col_widths(ws, {"A": 20, "B": 16, "C": 18, "D": 16, "E": 3, "F": 20, "G": 3, "H": 16, "I": 3, "J": 26, "K": 14})

header_cell(ws, 1, 1, "اسم الموظف")
header_cell(ws, 1, 2, "الراتب الأساسي")
header_cell(ws, 1, 3, "أجر ساعة الإضافي")
header_cell(ws, 1, 4, "بدل يوم العطلة")

employees = [
    ("إسماعيل", 3000, 15, 120),
    ("صادق", 2500, 12.5, 100),
    ("إسلام", 1800, 10, 72),
]
for i, (name, sal, ot, hol) in enumerate(employees, start=2):
    style_cell(ws, i, 1, editable=True, align="right").value = name
    style_cell(ws, i, 2, editable=True, number_format=CUR_FMT).value = sal
    style_cell(ws, i, 3, editable=True, number_format=CUR_FMT).value = ot
    style_cell(ws, i, 4, editable=True, number_format=CUR_FMT).value = hol

add_table(ws, "tblEmployees", f"A1:D{1 + len(employees)}", style="TableStyleLight1")

header_cell(ws, 1, 6, "التصنيف")
categories = ["ديزل", "مصاريف شخصية", "تجهيز سيارة", "صيانة", "قطع غيار", "اتصالات", "إعلانات", "سلفة عامل", "أخرى"]
for i, cat in enumerate(categories, start=2):
    style_cell(ws, i, 6, editable=True, align="right").value = cat
add_table(ws, "tblExpCat", f"F1:F{1 + len(categories)}", style="TableStyleLight1")

header_cell(ws, 1, 8, "طريقة الدفع")
methods = ["نقدًا", "تحويل بنكي", "بطاقة", "أخرى"]
for i, m in enumerate(methods, start=2):
    style_cell(ws, i, 8, editable=True, align="right").value = m
add_table(ws, "tblPayMethods", f"H1:H{1 + len(methods)}", style="TableStyleLight1")

header_cell(ws, 1, 10, "إعدادات الشركة")
ws.merge_cells("K1:K1")
style_cell(ws, 1, 11, editable=False, align="center").value = "القيمة"
ws.cell(row=1, column=11).font = f_header
ws.cell(row=1, column=11).fill = fill_header

style_cell(ws, 2, 10, editable=False, align="right").value = "وقت انتهاء الدوام الرسمي"
c = style_cell(ws, 2, 11, editable=True, number_format=TIME_FMT)
from datetime import time
c.value = time(23, 0, 0)

style_cell(ws, 3, 10, editable=False, align="right").value = "العملة"
style_cell(ws, 3, 11, editable=True, align="center").value = "ر.س"

ws.sheet_state = "hidden"

# Named ranges (Latin identifiers = internal plumbing, not user-visible content)
add_defined_name(wb, "WorkEndTime", f"'{SET}'!$K$2")
add_defined_name(wb, "EmployeeList",
                  f"OFFSET('{SET}'!$A$2,0,0,COUNTA('{SET}'!$A$2:$A$1000),1)")
add_defined_name(wb, "ExpenseCategoryList",
                  f"OFFSET('{SET}'!$F$2,0,0,COUNTA('{SET}'!$F$2:$F$1000),1)")
add_defined_name(wb, "PaymentMethodList",
                  f"OFFSET('{SET}'!$H$2,0,0,COUNTA('{SET}'!$H$2:$H$1000),1)")

# ============================================================================
# SALES SHEET
# ============================================================================
ws = ws_sales
set_col_widths(ws, {"A": 13, "B": 13, "C": 20, "D": 26, "E": 14, "F": 14, "G": 14})
ws.sheet_view.showGridLines = False

title_bar(ws, 1, 1, 7, "المبيعات")
ws.freeze_panes = "A2"

SALES_HEAD_ROW = 2
SALES_FIRST = 3
SALES_SEED_ROWS = 20
SALES_LAST = SALES_FIRST + SALES_SEED_ROWS - 1

headers = ["التاريخ", "اليوم", "العميل", "البيان", "المبلغ", "طريقة الدفع", "بداية الأسبوع"]
for j, h in enumerate(headers, start=1):
    header_cell(ws, SALES_HEAD_ROW, j, h)
ws.row_dimensions[SALES_HEAD_ROW].height = 32

for r in range(SALES_FIRST, SALES_LAST + 1):
    style_cell(ws, r, 1, editable=True, number_format=DATE_FMT)
    c = style_cell(ws, r, 2, editable=False)
    c.value = f'=IF(A{r}="","",CHOOSE(WEEKDAY(A{r},2),"الاثنين","الثلاثاء","الأربعاء","الخميس","الجمعة","السبت","الأحد"))'
    style_cell(ws, r, 3, editable=True, align="right")
    style_cell(ws, r, 4, editable=True, align="right")
    style_cell(ws, r, 5, editable=True, number_format=CUR_FMT)
    style_cell(ws, r, 6, editable=True)
    c = style_cell(ws, r, 7, editable=False, number_format=DATE_FMT)
    c.value = f'=IF(A{r}="","",A{r}-WEEKDAY(A{r},3))'

add_table(ws, "tblSales", f"A{SALES_HEAD_ROW}:G{SALES_LAST}", style="TableStyleLight9", bands=False)

dv_pay = list_dv("PaymentMethodList", "طريقة الدفع", "اختر طريقة الدفع من القائمة.")
ws.add_data_validation(dv_pay)
dv_pay.add(f"F{SALES_FIRST}:F{SALES_LAST}")

dv_amt = positive_number_dv("المبلغ")
ws.add_data_validation(dv_amt)
dv_amt.add(f"E{SALES_FIRST}:E{SALES_LAST}")

# ---- Totals strip ----
TOT_ROW = SALES_LAST + 2
title_bar(ws, TOT_ROW, 1, 7, "الإجمالي", fill=fill_section, font=f_section, height=24)
r = TOT_ROW + 1
style_cell(ws, r, 1, editable=False, align="right").value = "الإجمالي الكلي"
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
c = style_cell(ws, r, 3, editable=False, number_format=CUR_FMT)
c.value = f"=SUM(tblSales[المبلغ])"
c.font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)

style_cell(ws, r, 5, editable=False, align="right").value = "إجمالي الشهر الحالي"
ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
c = style_cell(ws, r, 7, editable=False, number_format=CUR_FMT)
c.value = (
    '=SUMIFS(tblSales[المبلغ],tblSales[التاريخ],">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),'
    'tblSales[التاريخ],"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))'
)
c.font = Font(name=FONT_NAME, size=12, bold=True, color=TEAL)
SALES_MONTH_TOTAL_CELL = f"'{SALES}'!G{r}"
SALES_GRAND_TOTAL_CELL = f"'{SALES}'!C{r}"

# ---- Weekly summary ----
WK_TITLE = r + 2
title_bar(ws, WK_TITLE, 1, 3, "ملخص المبيعات الأسبوعي (تلقائي)", fill=fill_section, font=f_section, height=24)
WK_HEAD = WK_TITLE + 1
header_cell(ws, WK_HEAD, 1, "بداية الأسبوع")
header_cell(ws, WK_HEAD, 2, "نهاية الأسبوع")
header_cell(ws, WK_HEAD, 3, "إجمالي المبيعات")
WK_FIRST = WK_HEAD + 1
WK_ROWS = 30
WK_LAST = WK_FIRST + WK_ROWS - 1

ws.cell(row=WK_FIRST, column=1,
        value=f'=IFERROR(SORT(UNIQUE(FILTER(tblSales[بداية الأسبوع],tblSales[بداية الأسبوع]<>""))),"")')
for r in range(WK_FIRST, WK_LAST + 1):
    style_cell(ws, r, 1, editable=False, number_format=DATE_FMT)
    c = style_cell(ws, r, 2, editable=False, number_format=DATE_FMT)
    c.value = f'=IF(A{r}="","",A{r}+6)'
    c = style_cell(ws, r, 3, editable=False, number_format=CUR_FMT)
    c.value = f'=IF(A{r}="","",SUMIFS(tblSales[المبلغ],tblSales[بداية الأسبوع],A{r}))'

ws.conditional_formatting.add(
    f"C{WK_FIRST}:C{WK_LAST}",
    DataBarRule(start_type="num", start_value=0, end_type="max", color=TEAL),
)

SALES_WEEK_RANGE = (WK_FIRST, WK_LAST)
print("Sales sheet built.")

# ============================================================================
# EXPENSES SHEET
# ============================================================================
ws = ws_exp
set_col_widths(ws, {"A": 13, "B": 17, "C": 26, "D": 14, "E": 24})
ws.sheet_view.showGridLines = False

title_bar(ws, 1, 1, 5, "المصروفات")
ws.freeze_panes = "A2"

EXP_HEAD_ROW = 2
EXP_FIRST = 3
EXP_SEED_ROWS = 20
EXP_LAST = EXP_FIRST + EXP_SEED_ROWS - 1

headers = ["التاريخ", "التصنيف", "البيان", "المبلغ", "ملاحظات"]
for j, h in enumerate(headers, start=1):
    header_cell(ws, EXP_HEAD_ROW, j, h)
ws.row_dimensions[EXP_HEAD_ROW].height = 28

for r in range(EXP_FIRST, EXP_LAST + 1):
    style_cell(ws, r, 1, editable=True, number_format=DATE_FMT)
    style_cell(ws, r, 2, editable=True, align="right")
    style_cell(ws, r, 3, editable=True, align="right")
    style_cell(ws, r, 4, editable=True, number_format=CUR_FMT)
    style_cell(ws, r, 5, editable=True, align="right")

add_table(ws, "tblExpenses", f"A{EXP_HEAD_ROW}:E{EXP_LAST}", style="TableStyleLight9", bands=False)

dv_cat = list_dv("ExpenseCategoryList", "التصنيف", "اختر تصنيف المصروف من القائمة.")
ws.add_data_validation(dv_cat)
dv_cat.add(f"B{EXP_FIRST}:B{EXP_LAST}")

dv_amt = positive_number_dv("المبلغ")
ws.add_data_validation(dv_amt)
dv_amt.add(f"D{EXP_FIRST}:D{EXP_LAST}")

# ---- Category totals ----
CAT_TITLE = EXP_LAST + 2
title_bar(ws, CAT_TITLE, 1, 2, "الإجمالي حسب التصنيف (تلقائي)", fill=fill_section, font=f_section, height=24)
CAT_HEAD = CAT_TITLE + 1
header_cell(ws, CAT_HEAD, 1, "التصنيف")
header_cell(ws, CAT_HEAD, 2, "الإجمالي")
CAT_FIRST = CAT_HEAD + 1
CAT_LAST = CAT_FIRST + len(categories) - 1

ws.cell(row=CAT_FIRST, column=1, value="=tblExpCat[التصنيف]")
for r in range(CAT_FIRST, CAT_LAST + 1):
    style_cell(ws, r, 1, editable=False, align="right")
    c = style_cell(ws, r, 2, editable=False, number_format=CUR_FMT)
    c.value = f'=IF(A{r}="","",SUMIFS(tblExpenses[المبلغ],tblExpenses[التصنيف],A{r}))'

ws.conditional_formatting.add(
    f"B{CAT_FIRST}:B{CAT_LAST}",
    DataBarRule(start_type="num", start_value=0, end_type="max", color=GOLD),
)

TOT_ROW = CAT_LAST + 2
style_cell(ws, TOT_ROW, 1, editable=False, align="right").value = "الإجمالي الكلي"
c = style_cell(ws, TOT_ROW, 2, editable=False, number_format=CUR_FMT)
c.value = "=SUM(tblExpenses[المبلغ])"
c.font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
EXP_GRAND_TOTAL_CELL = f"'{EXP}'!B{TOT_ROW}"

r2 = TOT_ROW + 1
style_cell(ws, r2, 1, editable=False, align="right").value = "إجمالي الشهر الحالي"
c = style_cell(ws, r2, 2, editable=False, number_format=CUR_FMT)
c.value = (
    '=SUMIFS(tblExpenses[المبلغ],tblExpenses[التاريخ],">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),'
    'tblExpenses[التاريخ],"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))'
)
c.font = Font(name=FONT_NAME, size=12, bold=True, color=RED)
EXP_MONTH_TOTAL_CELL = f"'{EXP}'!B{r2}"

EXP_CAT_RANGE = (CAT_FIRST, CAT_LAST)
print("Expenses sheet built.")

# ============================================================================
# EMPLOYEES SHEET  (Payroll summary [dynamic array] + Work Log + Transactions)
# ============================================================================
ws = ws_emp
set_col_widths(ws, {
    "A": 16, "B": 15, "C": 16, "D": 15, "E": 14,
    "F": 14, "G": 13, "H": 13, "I": 13, "J": 15,
})
ws.sheet_view.showGridLines = False

title_bar(ws, 1, 1, 10, "الموظفون")
note(ws, 2, 1, 10,
     "ملاحظة: جدول ملخص الرواتب أدناه يُحسب تلقائيًا بالكامل من الإعدادات وسجل العمل والحركات - لا تقم بالإدخال فيه مباشرة.")
ws.freeze_panes = "A2"

# ---- Section 1: Payroll summary (dynamic array, auto-grows with employees) ----
SEC1_TITLE = 3
title_bar(ws, SEC1_TITLE, 1, 10, "أولاً: ملخص الرواتب (تلقائي بالكامل)", fill=fill_section, font=f_section)

PR_HEAD = SEC1_TITLE + 1
pr_headers = ["اسم الموظف", "الراتب الأساسي", "إجمالي ساعات الإضافي", "قيمة الإضافي",
              "أيام العمل بالعطلة", "بدل العطلة", "السلف", "الخصومات", "المكافآت", "صافي الراتب"]
for j, h in enumerate(pr_headers, start=1):
    header_cell(ws, PR_HEAD, j, h)
ws.row_dimensions[PR_HEAD].height = 34

PR_FIRST = PR_HEAD + 1
PR_BAND_ROWS = 20
PR_LAST = PR_FIRST + PR_BAND_ROWS - 1
CUR_COLS = {2, 4, 6, 7, 8, 9, 10}
for r in range(PR_FIRST, PR_LAST + 1):
    for col in range(1, 11):
        align = "right" if col == 1 else "center"
        fmt = CUR_FMT if col in CUR_COLS else (HOURS_FMT if col == 3 else ("0" if col == 5 else None))
        style_cell(ws, r, col, editable=False, number_format=fmt, align=align)

R = PR_FIRST
ws.cell(row=R, column=1, value="=tblEmployees[اسم الموظف]")
ws.cell(row=R, column=2,
        value=f'=IFERROR(XLOOKUP(A{R}#,tblEmployees[اسم الموظف],tblEmployees[الراتب الأساسي]),"")')
ws.cell(row=R, column=3,
        value=f'=IFERROR(SUMIFS(tblWorkLog[ساعات الإضافي],tblWorkLog[اسم الموظف],A{R}#),0)')
ws.cell(row=R, column=4,
        value=f'=IFERROR(SUMIFS(tblWorkLog[قيمة الإضافي],tblWorkLog[اسم الموظف],A{R}#),0)')
ws.cell(row=R, column=5,
        value=f'=IFERROR(COUNTIFS(tblWorkLog[اسم الموظف],A{R}#,tblWorkLog[يوم عطلة؟],"نعم"),0)')
ws.cell(row=R, column=6,
        value=f'=IFERROR(E{R}#*XLOOKUP(A{R}#,tblEmployees[اسم الموظف],tblEmployees[بدل يوم العطلة]),0)')
ws.cell(row=R, column=7,
        value=f'=IFERROR(SUMIFS(tblTransactions[المبلغ],tblTransactions[اسم الموظف],A{R}#,tblTransactions[نوع العملية],"سلفة"),0)')
ws.cell(row=R, column=8,
        value=f'=IFERROR(SUMIFS(tblTransactions[المبلغ],tblTransactions[اسم الموظف],A{R}#,tblTransactions[نوع العملية],"خصم"),0)')
ws.cell(row=R, column=9,
        value=f'=IFERROR(SUMIFS(tblTransactions[المبلغ],tblTransactions[اسم الموظف],A{R}#,tblTransactions[نوع العملية],"مكافأة"),0)')
ws.cell(row=R, column=10, value=f'=IF(A{R}#="","",B{R}#+D{R}#+F{R}#+I{R}#-G{R}#-H{R}#)')

ws.auto_filter.ref = f"A{PR_HEAD}:J{PR_LAST}"

# zebra banding only on populated rows
ws.conditional_formatting.add(
    f"A{PR_FIRST}:J{PR_LAST}",
    FormulaRule(formula=[f'AND($A{PR_FIRST}<>"",MOD(ROW(),2)=0)'], fill=fill_band),
)

PR_NET_RANGE = f"J{PR_FIRST}#"
PR_OT_RANGE = f"D{PR_FIRST}#"
PR_NAME_RANGE = f"A{PR_FIRST}#"

# ---- Section 2: Daily work log ----
SEC2_TITLE = PR_LAST + 2
title_bar(ws, SEC2_TITLE, 1, 9, "ثانيًا: سجل العمل اليومي", fill=fill_section, font=f_section)

WL_HEAD = SEC2_TITLE + 1
wl_headers = ["التاريخ", "اسم الموظف", "وقت الانصراف", "وقت انتهاء العمل", "ساعات الإضافي",
              "قيمة الإضافي", "يوم عطلة؟", "المشروع / الموقع", "ملاحظات"]
for j, h in enumerate(wl_headers, start=1):
    header_cell(ws, WL_HEAD, j, h)
ws.row_dimensions[WL_HEAD].height = 34

WL_FIRST = WL_HEAD + 1
WL_SEED_ROWS = 15
WL_LAST = WL_FIRST + WL_SEED_ROWS - 1

for r in range(WL_FIRST, WL_LAST + 1):
    style_cell(ws, r, 1, editable=True, number_format=DATE_FMT)
    style_cell(ws, r, 2, editable=True, align="right")
    style_cell(ws, r, 3, editable=True, number_format=TIME_FMT)
    style_cell(ws, r, 4, editable=True, number_format=TIME_FMT)
    c = style_cell(ws, r, 5, editable=False, number_format=HOURS_FMT)
    c.value = f'=IF(OR(A{r}="",D{r}=""),"",MAX(0,(D{r}+IF(D{r}<TIME(12,0,0),1,0)-WorkEndTime))*24)'
    c = style_cell(ws, r, 6, editable=False, number_format=CUR_FMT)
    c.value = f'=IF(E{r}="","",E{r}*IFERROR(XLOOKUP(B{r},tblEmployees[اسم الموظف],tblEmployees[أجر ساعة الإضافي]),0))'
    style_cell(ws, r, 7, editable=True)
    style_cell(ws, r, 8, editable=True, align="right")
    style_cell(ws, r, 9, editable=True, align="right")

add_table(ws, "tblWorkLog", f"A{WL_HEAD}:I{WL_LAST}", style="TableStyleLight9", bands=False)

dv_emp1 = list_dv("EmployeeList", "اسم الموظف", "اختر اسم الموظف من الإعدادات.")
ws.add_data_validation(dv_emp1)
dv_emp1.add(f"B{WL_FIRST}:B{WL_LAST}")

dv_holiday = list_dv('"نعم,لا"', "يوم عطلة؟", "هل هذا اليوم يوم عطلة رسمية؟")
ws.add_data_validation(dv_holiday)
dv_holiday.add(f"G{WL_FIRST}:G{WL_LAST}")

ws.conditional_formatting.add(
    f"E{WL_FIRST}:E{WL_LAST}",
    CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor="FDE9C8")),
)

# ---- Section 3: Employee transactions (advances / deductions / bonuses) ----
SEC3_TITLE = WL_LAST + 2
title_bar(ws, SEC3_TITLE, 1, 6, "ثالثًا: حركات الموظفين (سلف / خصومات / مكافآت)", fill=fill_section, font=f_section)

TR_HEAD = SEC3_TITLE + 1
tr_headers = ["التاريخ", "اسم الموظف", "نوع العملية", "المبلغ", "السبب", "ملاحظات"]
for j, h in enumerate(tr_headers, start=1):
    header_cell(ws, TR_HEAD, j, h)
ws.row_dimensions[TR_HEAD].height = 28

TR_FIRST = TR_HEAD + 1
TR_SEED_ROWS = 15
TR_LAST = TR_FIRST + TR_SEED_ROWS - 1

for r in range(TR_FIRST, TR_LAST + 1):
    style_cell(ws, r, 1, editable=True, number_format=DATE_FMT)
    style_cell(ws, r, 2, editable=True, align="right")
    style_cell(ws, r, 3, editable=True)
    style_cell(ws, r, 4, editable=True, number_format=CUR_FMT)
    style_cell(ws, r, 5, editable=True, align="right")
    style_cell(ws, r, 6, editable=True, align="right")

add_table(ws, "tblTransactions", f"A{TR_HEAD}:F{TR_LAST}", style="TableStyleLight9", bands=False)

dv_emp2 = list_dv("EmployeeList", "اسم الموظف", "اختر اسم الموظف من الإعدادات.")
ws.add_data_validation(dv_emp2)
dv_emp2.add(f"B{TR_FIRST}:B{TR_LAST}")

dv_type = list_dv('"سلفة,خصم,مكافأة"', "نوع العملية", "اختر نوع العملية.")
ws.add_data_validation(dv_type)
dv_type.add(f"C{TR_FIRST}:C{TR_LAST}")

dv_amt3 = positive_number_dv("المبلغ")
ws.add_data_validation(dv_amt3)
dv_amt3.add(f"D{TR_FIRST}:D{TR_LAST}")

print("Employees sheet built.")

# ============================================================================
# DASHBOARD SHEET
# ============================================================================
ws = ws_dash
set_col_widths(ws, {c: 11 for c in "ABCDEFGHIJKL"})
ws.sheet_view.showGridLines = False

title_bar(ws, 1, 1, 12, "لوحة التحكم - نظرة عامة على الأعمال")
ws.freeze_panes = "A2"
note(ws, 2, 1, 12, '=" تحديث تلقائي بتاريخ "&TEXT(TODAY(),"yyyy/mm/dd")')

KPI_LABEL_ROW = 3
KPI_VALUE_ROW = 4
ws.row_dimensions[KPI_VALUE_ROW].height = 34

cards = [
    ("المبيعات الشهرية", SALES_MONTH_TOTAL_CELL, TEAL, CUR_FMT),
    ("المصروفات الشهرية", EXP_MONTH_TOTAL_CELL, RED, CUR_FMT),
    ("صافي الربح", None, NAVY, CUR_FMT),
    ("إجمالي الرواتب", None, "5B4B8A", CUR_FMT),
    ("إجمالي قيمة الإضافي", None, GOLD, CUR_FMT),
    ("عدد الموظفين", None, "3D6E4F", "0"),
]

kpi_value_cells = {}
for i, (label, formula, color, fmt) in enumerate(cards):
    col = 1 + i * 2
    ws.merge_cells(start_row=KPI_LABEL_ROW, start_column=col, end_row=KPI_LABEL_ROW, end_column=col + 1)
    lc = ws.cell(row=KPI_LABEL_ROW, column=col, value=label)
    lc.font = f_kpi_label
    lc.fill = PatternFill("solid", fgColor=color)
    lc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for cc in range(col, col + 2):
        ws.cell(row=KPI_LABEL_ROW, column=cc).fill = PatternFill("solid", fgColor=color)
        ws.cell(row=KPI_LABEL_ROW, column=cc).protection = LOCKED

    ws.merge_cells(start_row=KPI_VALUE_ROW, start_column=col, end_row=KPI_VALUE_ROW, end_column=col + 1)
    vc = ws.cell(row=KPI_VALUE_ROW, column=col)
    vc.font = f_kpi_value
    vc.fill = PatternFill("solid", fgColor=color)
    vc.number_format = fmt
    vc.alignment = Alignment(horizontal="center", vertical="center")
    vc.protection = LOCKED
    if formula:
        vc.value = f"={formula}"
    for cc in range(col, col + 2):
        ws.cell(row=KPI_VALUE_ROW, column=cc).fill = PatternFill("solid", fgColor=color)
    kpi_value_cells[label] = ws.cell(row=KPI_VALUE_ROW, column=col).coordinate

# formulas that depend on other card cells, filled in after all coords are known
ws[kpi_value_cells["صافي الربح"]] = f"={kpi_value_cells['المبيعات الشهرية']}-{kpi_value_cells['المصروفات الشهرية']}"
ws[kpi_value_cells["إجمالي الرواتب"]] = f"=SUM('{EMP}'!{PR_NET_RANGE})"
ws[kpi_value_cells["إجمالي قيمة الإضافي"]] = f"=SUM('{EMP}'!{PR_OT_RANGE})"
ws[kpi_value_cells["عدد الموظفين"]] = "=COUNTA(tblEmployees[اسم الموظف])"

# lighten kpi fonts for readability (dark backgrounds -> white text already set)
ws.conditional_formatting.add(
    kpi_value_cells["صافي الربح"],
    CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor=RED)),
)

# ---- hidden helper row feeding the Sales/Expenses/Profit chart ----
HELP_ROW1, HELP_ROW2 = 6, 7
ws.cell(row=HELP_ROW1, column=1, value="المبيعات")
ws.cell(row=HELP_ROW1, column=2, value="المصروفات")
ws.cell(row=HELP_ROW1, column=3, value="صافي الربح")
ws.cell(row=HELP_ROW2, column=1, value=f"={kpi_value_cells['المبيعات الشهرية']}")
ws.cell(row=HELP_ROW2, column=2, value=f"={kpi_value_cells['المصروفات الشهرية']}")
ws.cell(row=HELP_ROW2, column=3, value=f"={kpi_value_cells['صافي الربح']}")
for row in (HELP_ROW1, HELP_ROW2):
    for col in range(1, 4):
        ws.cell(row=row, column=col).font = f_note
ws.row_dimensions[HELP_ROW1].hidden = True
ws.row_dimensions[HELP_ROW2].hidden = True

# ---- Charts ----
chart1 = BarChart()
chart1.type = "col"
chart1.title = "المبيعات والمصروفات وصافي الربح (الشهر الحالي)"
chart1.style = 10
chart1.y_axis.title = "ر.س"
chart1.visible_cells_only = False
data_ref = Reference(ws, min_col=1, max_col=3, min_row=HELP_ROW2, max_row=HELP_ROW2)
cats_ref = Reference(ws, min_col=1, max_col=3, min_row=HELP_ROW1, max_row=HELP_ROW1)
chart1.add_data(data_ref, titles_from_data=False, from_rows=True)
set_text_categories(chart1, cats_ref)
chart1.legend = None
chart1.width = 15
chart1.height = 8.5
ws.add_chart(chart1, "A9")

chart2 = PieChart()
chart2.title = "توزيع المصروفات حسب التصنيف"
chart2.style = 10
cat_ref2 = Reference(ws_exp, min_col=1, max_col=1, min_row=EXP_CAT_RANGE[0], max_row=EXP_CAT_RANGE[1])
data_ref2 = Reference(ws_exp, min_col=2, max_col=2, min_row=EXP_CAT_RANGE[0], max_row=EXP_CAT_RANGE[1])
chart2.add_data(data_ref2, titles_from_data=False)
set_text_categories(chart2, cat_ref2)
chart2.dataLabels = DataLabelList()
chart2.dataLabels.showPercent = True
chart2.width = 15
chart2.height = 8.5
ws.add_chart(chart2, "G9")

chart3 = LineChart()
chart3.title = "اتجاه المبيعات الأسبوعي"
chart3.style = 12
chart3.y_axis.title = "ر.س"
chart3.visible_cells_only = False
cat_ref3 = Reference(ws_sales, min_col=1, max_col=1, min_row=SALES_WEEK_RANGE[0], max_row=SALES_WEEK_RANGE[1])
data_ref3 = Reference(ws_sales, min_col=3, max_col=3, min_row=SALES_WEEK_RANGE[0], max_row=SALES_WEEK_RANGE[1])
chart3.add_data(data_ref3, titles_from_data=False)
chart3.set_categories(cat_ref3)
chart3.legend = None
chart3.width = 30
chart3.height = 9
ws.add_chart(chart3, "A26")

print("Dashboard sheet built.")

# ============================================================================
# PROTECTION, TAB COLORS, FINAL TOUCHES
# ============================================================================
tab_colors = {ws_dash: NAVY, ws_sales: TEAL, ws_exp: RED, ws_emp: GOLD, ws_set: "999999"}
for ws_, color in tab_colors.items():
    ws_.sheet_properties.tabColor = color

for ws_ in (ws_dash, ws_sales, ws_exp, ws_emp, ws_set):
    ws_.protection.sheet = True
    ws_.protection.autoFilter = False
    ws_.protection.sort = False
    ws_.protection.formatCells = False
    ws_.protection.formatColumns = False
    ws_.protection.formatRows = False
    ws_.protection.selectLockedCells = False
    ws_.protection.selectUnlockedCells = False

wb.active = 0
wb._sheets = [ws_dash, ws_sales, ws_exp, ws_emp, ws_set]

OUT_PATH = "business_workbook.xlsx"
wb.save(OUT_PATH)
print(f"Workbook saved to {OUT_PATH}")
