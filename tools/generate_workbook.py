# -*- coding: utf-8 -*-
"""
مولد ملف إكسل - نظام إدارة شركة المقاولات
Generator script for the Construction Company Management Workbook.
"""
import re
import datetime
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.styles.protection import Protection
from openpyxl.formatting.rule import CellIsRule, FormulaRule, DataBarRule
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.chart import BarChart, LineChart, PieChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.colors import ColorChoice

# ----------------------------------------------------------------------------
# الثوابت العامة - Global constants (palette, fonts, sizes)
# ----------------------------------------------------------------------------
FONT_NAME = "Arial"

C_PRIMARY = "1F4E78"      # navy - headers / titles
C_PRIMARY2 = "2E75B6"     # medium blue - sub headers
C_ACCENT = "D9E6F2"       # light blue - card / band background
C_SUCCESS = "1E7145"      # green - positive
C_SUCCESS_L = "E2F0D9"
C_DANGER = "B02A2A"       # red - negative
C_DANGER_L = "FBE2E2"
C_WARNING = "B25B00"      # orange
C_WARNING_L = "FCEBD7"
C_INPUT = "FFF6D5"        # highlighted input cells
C_LOCKED = "F2F2F2"       # locked/formula cells
C_WHITE = "FFFFFF"
C_TEXT = "1F1F1F"
C_GRAY_TEXT = "6B6B6B"
C_BORDER = "BFBFBF"

PROTECT_PASSWORD = "2026$Construct"

FMT_MONEY = '#,##0.00'
FMT_INT = '#,##0'
FMT_DATE = 'yyyy-mm-dd'
FMT_TIME = 'hh:mm'
FMT_PCT = '0.0%'

thin = Side(style="thin", color=C_BORDER)
BORDER_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)

FUNCS_XLWS = ["SORTBY", "SORT", "UNIQUE", "FILTER", "SEQUENCE", "RANDARRAY"]
FUNCS_PLAIN = ["XLOOKUP", "XMATCH", "LET", "IFS", "SWITCH", "TEXTJOIN",
               "MAXIFS", "MINIFS", "ISOWEEKNUM"]


def xf(formula: str) -> str:
    """Prefix modern Excel functions with the _xlfn/_xlws tokens required
    inside the raw OOXML formula string so Excel recognises them."""
    for name in FUNCS_XLWS:
        formula = re.sub(rf'(?<![A-Za-z0-9_.]){name}\(',
                          f'_xlfn._xlws.{name}(', formula)
    for name in FUNCS_PLAIN:
        formula = re.sub(rf'(?<![A-Za-z0-9_.]){name}\(',
                          f'_xlfn.{name}(', formula)
    return formula


wb = Workbook()
wb.remove(wb.active)
wb.calculation = CalcProperties(fullCalcOnLoad=True)
wb.properties.title = "نظام إدارة شركة المقاولات"
wb.properties.creator = "Claude - نظام الإدارة الآلي"

# ----------------------------------------------------------------------------
# دوال مساعدة للتنسيق - styling helpers
# ----------------------------------------------------------------------------

def new_sheet(name, hidden=False):
    ws = wb.create_sheet(name)
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    if hidden:
        ws.sheet_state = "hidden"
    return ws


def set_col_widths(ws, widths: dict):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def big_title(ws, cell_range, text_or_formula, size=18, fill=C_PRIMARY,
              font_color=C_WHITE, height=32):
    ws.merge_cells(cell_range)
    top_left = cell_range.split(":")[0]
    c = ws[top_left]
    c.value = text_or_formula
    c.font = Font(name=FONT_NAME, size=size, bold=True, color=font_color)
    c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(horizontal="center", vertical="center")
    row = int(re.search(r"\d+", top_left).group())
    ws.row_dimensions[row].height = height
    # paint the whole merged band
    for col_cell in ws[cell_range][0]:
        col_cell.fill = PatternFill("solid", fgColor=fill)


def sub_title(ws, cell_range, text_or_formula, size=10, italic=True,
              color=C_GRAY_TEXT, align="center"):
    ws.merge_cells(cell_range)
    top_left = cell_range.split(":")[0]
    c = ws[top_left]
    c.value = text_or_formula
    c.font = Font(name=FONT_NAME, size=size, italic=italic, color=color)
    c.alignment = Alignment(horizontal=align, vertical="center")


def block_title(ws, cell_range, text, fill=C_PRIMARY2, size=11):
    ws.merge_cells(cell_range)
    top_left = cell_range.split(":")[0]
    c = ws[top_left]
    c.value = text
    c.font = Font(name=FONT_NAME, size=size, bold=True, color=C_WHITE)
    c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(horizontal="center", vertical="center")
    for col_cell in ws[cell_range][0]:
        col_cell.fill = PatternFill("solid", fgColor=fill)


def header_cell(ws, ref, text, fill=C_PRIMARY):
    c = ws[ref]
    c.value = text
    c.font = Font(name=FONT_NAME, size=10, bold=True, color=C_WHITE)
    c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER_ALL


def input_cell(ws, ref, value=None, number_format=None, align="center"):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.fill = PatternFill("solid", fgColor=C_INPUT)
    c.font = Font(name=FONT_NAME, size=10, color=C_TEXT)
    c.alignment = Alignment(horizontal=align, vertical="center")
    c.border = BORDER_ALL
    c.protection = Protection(locked=False)
    if number_format:
        c.number_format = number_format
    return c


def formula_cell(ws, ref, formula, number_format=None, align="center",
                  bold=False, color=C_TEXT):
    c = ws[ref]
    c.value = xf(formula) if isinstance(formula, str) and formula.startswith("=") else formula
    c.fill = PatternFill("solid", fgColor=C_LOCKED)
    c.font = Font(name=FONT_NAME, size=10, color=color, bold=bold)
    c.alignment = Alignment(horizontal=align, vertical="center")
    c.border = BORDER_ALL
    c.protection = Protection(locked=True)
    if number_format:
        c.number_format = number_format
    return c


def plain_cell(ws, ref, value, size=10, bold=False, color=C_TEXT,
               align="right", number_format=None, fill=None):
    c = ws[ref]
    c.value = xf(value) if isinstance(value, str) and value.startswith("=") else value
    c.font = Font(name=FONT_NAME, size=size, bold=bold, color=color)
    c.alignment = Alignment(horizontal=align, vertical="center")
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if number_format:
        c.number_format = number_format
    return c


def add_table(ws, name, ref, style="TableStyleMedium9", totals=False):
    tab = Table(displayName=name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(
        name=style, showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False)
    ws.add_table(tab)
    return tab


def add_name(name, ref_formula):
    if ref_formula.startswith("="):
        ref_formula = ref_formula[1:]
    wb.defined_names[name] = DefinedName(name, attr_text=ref_formula)


def add_list_dv(ws, cell_range, source, allow_blank=True):
    dv = DataValidation(type="list", formula1=source, allow_blank=allow_blank,
                         showDropDown=False, showErrorMessage=True,
                         errorTitle="قيمة غير صحيحة", error="الرجاء الاختيار من القائمة المنسدلة")
    ws.add_data_validation(dv)
    dv.add(cell_range)
    return dv


print("setup ok")

# ============================================================================
# ورقة الإعدادات (مخفية) - قاعدة البيانات الرئيسية للنظام
# ============================================================================
ws_set = new_sheet("الإعدادات", hidden=True)
set_col_widths(ws_set, {"A": 20, "B": 16, "C": 20, "D": 16, "E": 12,
                         "F": 3, "G": 22, "H": 3, "I": 18, "J": 3,
                         "K": 32, "L": 20})

big_title(ws_set, "A1:L1",
          "قاعدة البيانات الرئيسية للنظام - لا تحذف أو تعيد تسمية هذه الورقة",
          size=13, height=26)
sub_title(ws_set, "A2:L2",
          "هذه الورقة هي المصدر الوحيد للبيانات الأساسية. أي تعديل هنا (إضافة موظف، فئة مصروف، ...) "
          "ينعكس تلقائياً على كل صفحات النظام دون تعديل أي معادلة.")

# ---- جدول الموظفين ----
EMP_HEAD_ROW = 3
EMP_FIRST = 4
EMP_SEED = [
    ("إسماعيل", 3000, 15, 120, "نشط"),
    ("صادق", 2500, 12.5, 100, "نشط"),
    ("إسلام", 1800, 10, 72, "نشط"),
]
EMP_BLANK_ROWS = 12
EMP_LAST = EMP_FIRST + len(EMP_SEED) + EMP_BLANK_ROWS - 1

for ref, text in zip("ABCDE", ["اسم الموظف", "الراتب الأساسي", "معدل الساعة الإضافية",
                                "أجر يوم العطلة", "الحالة"]):
    header_cell(ws_set, f"{ref}{EMP_HEAD_ROW}", text)

r = EMP_FIRST
for name, salary, ot_rate, hol_rate, status in EMP_SEED:
    input_cell(ws_set, f"A{r}", name)
    input_cell(ws_set, f"B{r}", salary, FMT_MONEY)
    input_cell(ws_set, f"C{r}", ot_rate, FMT_MONEY)
    input_cell(ws_set, f"D{r}", hol_rate, FMT_MONEY)
    input_cell(ws_set, f"E{r}", status)
    r += 1
for _ in range(EMP_BLANK_ROWS):
    for col, fmt in [("A", None), ("B", FMT_MONEY), ("C", FMT_MONEY), ("D", FMT_MONEY), ("E", None)]:
        input_cell(ws_set, f"{col}{r}", None, fmt)
    r += 1

add_table(ws_set, "جدول_الموظفين", f"A{EMP_HEAD_ROW}:E{EMP_LAST}")
add_list_dv(ws_set, f"E{EMP_FIRST}:E{EMP_LAST}", '"نشط,غير نشط"')

# ---- جدول فئات المصروفات ----
CAT_HEAD_ROW = 3
CAT_FIRST = 4
CATEGORIES = ["ديزل", "مصاريف شخصية", "تجهيز مركبة جديدة", "اتصالات", "إعلانات",
              "صيانة المركبات", "قطع غيار", "معدات", "إيجار", "رسوم حكومية",
              "ضيافة", "أخرى"]
CAT_BLANK_ROWS = 6
CAT_LAST = CAT_FIRST + len(CATEGORIES) + CAT_BLANK_ROWS - 1

header_cell(ws_set, f"G{CAT_HEAD_ROW}", "الفئة")
r = CAT_FIRST
for cat in CATEGORIES:
    input_cell(ws_set, f"G{r}", cat)
    r += 1
for _ in range(CAT_BLANK_ROWS):
    input_cell(ws_set, f"G{r}")
    r += 1
add_table(ws_set, "جدول_فئات_المصروفات", f"G{CAT_HEAD_ROW}:G{CAT_LAST}")

# ---- جدول طرق الدفع ----
PM_HEAD_ROW = 3
PM_FIRST = 4
PAYMENT_METHODS = ["نقدي", "تحويل بنكي", "بطاقة", "أخرى"]
PM_BLANK_ROWS = 4
PM_LAST = PM_FIRST + len(PAYMENT_METHODS) + PM_BLANK_ROWS - 1

header_cell(ws_set, f"I{PM_HEAD_ROW}", "طريقة الدفع")
r = PM_FIRST
for pm in PAYMENT_METHODS:
    input_cell(ws_set, f"I{r}", pm)
    r += 1
for _ in range(PM_BLANK_ROWS):
    input_cell(ws_set, f"I{r}")
    r += 1
add_table(ws_set, "جدول_طرق_الدفع", f"I{PM_HEAD_ROW}:I{PM_LAST}")

# ---- إعدادات الشركة (مفتاح / قيمة) ----
CS_HEAD_ROW = 3
CS_FIRST = 4
header_cell(ws_set, f"K{CS_HEAD_ROW}", "الإعداد")
header_cell(ws_set, f"L{CS_HEAD_ROW}", "القيمة")

CURRENT_MONTH = datetime.date(2026, 7, 1)
company_rows = [
    ("اسم الشركة", "شركة الأفق للمقاولات العامة", None),
    ("الفترة الحالية لإعداد الرواتب والتقارير", CURRENT_MONTH, FMT_DATE),
    ("السنة المالية", 2026, "0"),
    ("العملة", "ريال سعودي", None),
    ("رمز العملة", "ر.س", None),
    ("وقت انتهاء الدوام الرسمي", datetime.time(23, 0), FMT_TIME),
    ("حد اعتبار الاستمرار بعد منتصف الليل", datetime.time(6, 0), FMT_TIME),
]
CS_LAST = CS_FIRST + len(company_rows) - 1
r = CS_FIRST
for label, value, fmt in company_rows:
    input_cell(ws_set, f"K{r}", label, align="right")
    input_cell(ws_set, f"L{r}", value, fmt)
    r += 1
add_table(ws_set, "جدول_إعدادات_الشركة", f"K{CS_HEAD_ROW}:L{CS_LAST}")

# ---- أسماء معرّفة (Named Ranges) للإعدادات الفردية ----
add_name("اسم_الشركة", f"الإعدادات!$L${CS_FIRST}")
add_name("الشهر_الحالي", f"الإعدادات!$L${CS_FIRST+1}")
add_name("السنة_المالية", f"الإعدادات!$L${CS_FIRST+2}")
add_name("العملة", f"الإعدادات!$L${CS_FIRST+3}")
add_name("رمز_العملة", f"الإعدادات!$L${CS_FIRST+4}")
add_name("وقت_نهاية_الدوام_الرسمي", f"الإعدادات!$L${CS_FIRST+5}")
add_name("حد_اعتبار_الاستمرار", f"الإعدادات!$L${CS_FIRST+6}")

# ---- قوائم منسدلة ديناميكية (تُستخدم في كل الصفحات) ----
add_name("قائمة_الموظفين", "جدول_الموظفين[اسم الموظف]")
add_name("قائمة_الموظفين_النشطين",
         xf('=IFERROR(FILTER(جدول_الموظفين[اسم الموظف],جدول_الموظفين[الحالة]="نشط"),"")'))
add_name("قائمة_فئات_المصروفات", "جدول_فئات_المصروفات[الفئة]")
add_name("قائمة_طرق_الدفع", "جدول_طرق_الدفع[طريقة الدفع]")

ws_set.freeze_panes = "A4"
print("settings sheet ok")

# ============================================================================
# ورقة المبيعات
# ============================================================================
ws_sal = new_sheet("المبيعات")
set_col_widths(ws_sal, {"A": 12, "B": 11, "C": 22, "D": 13, "E": 24, "F": 14,
                         "G": 14, "H": 16, "I": 10, "J": 3,
                         "K": 16, "L": 13, "M": 3, "N": 12, "O": 13, "P": 3,
                         "Q": 22, "R": 13, "S": 3, "T": 9, "U": 22, "V": 13})

big_title(ws_sal, "A1:I1", '="سجل المبيعات  -  " & اسم_الشركة', size=16)
sub_title(ws_sal, "A2:I2",
          '="الفترة الحالية: " & TEXT(الشهر_الحالي,"yyyy-mm") & "   |   العملة: " & العملة')

SAL_HEAD_ROW = 4
SAL_FIRST = 5
SAL_HEADERS = ["التاريخ", "اليوم", "العميل", "رقم الفاتورة", "البيان", "المبلغ",
               "طريقة الدفع", "الأسبوع", "الشهر"]
for i, h in enumerate(SAL_HEADERS):
    header_cell(ws_sal, f"{get_column_letter(1+i)}{SAL_HEAD_ROW}", h)

SAL_SAMPLE = [
    (datetime.date(2026, 5, 20), "شركة الإعمار المتحدة", "INV-0932", "توريد حديد تسليح", 30000, "نقدي"),
    (datetime.date(2026, 6, 15), "شركة الرياض للمقاولات", "INV-0987", "أعمال صيانة عامة", 12000, "تحويل بنكي"),
    (datetime.date(2026, 7, 2), "مؤسسة البناء الحديث", "INV-1001", "توريد مواد بناء", 15000, "نقدي"),
    (datetime.date(2026, 7, 5), "شركة الإعمار المتحدة", "INV-1002", "أعمال حفر وردم", 22000, "تحويل بنكي"),
    (datetime.date(2026, 7, 10), "مؤسسة البناء الحديث", "INV-1003", "أعمال دهانات", 8500, "بطاقة"),
]
SAL_BLANK_ROWS = 15
SAL_LAST = SAL_FIRST + len(SAL_SAMPLE) + SAL_BLANK_ROWS - 1


def sales_row_formulas(ws, r):
    formula_cell(ws, f"B{r}", f'=IF([@التاريخ]="","",CHOOSE(WEEKDAY([@التاريخ],2),'
                               f'"الاثنين","الثلاثاء","الأربعاء","الخميس","الجمعة","السبت","الأحد"))')
    formula_cell(ws, f"H{r}", f'=IF([@التاريخ]="","","أسبوع "&TEXT(ISOWEEKNUM([@التاريخ]),"00")&'
                               f'" - "&YEAR([@التاريخ]))')
    formula_cell(ws, f"I{r}", '=IF([@التاريخ]="","",TEXT([@التاريخ],"yyyy-mm"))')


r = SAL_FIRST
for d, cust, inv, desc, amt, pm in SAL_SAMPLE:
    input_cell(ws_sal, f"A{r}", d, FMT_DATE)
    input_cell(ws_sal, f"C{r}", cust, align="right")
    input_cell(ws_sal, f"D{r}", inv)
    input_cell(ws_sal, f"E{r}", desc, align="right")
    input_cell(ws_sal, f"F{r}", amt, FMT_MONEY)
    input_cell(ws_sal, f"G{r}", pm)
    sales_row_formulas(ws_sal, r)
    r += 1
for _ in range(SAL_BLANK_ROWS):
    input_cell(ws_sal, f"A{r}", None, FMT_DATE)
    input_cell(ws_sal, f"C{r}", None, align="right")
    input_cell(ws_sal, f"D{r}")
    input_cell(ws_sal, f"E{r}", None, align="right")
    input_cell(ws_sal, f"F{r}", None, FMT_MONEY)
    input_cell(ws_sal, f"G{r}")
    sales_row_formulas(ws_sal, r)
    r += 1

SAL_TOTALS_ROW = SAL_LAST + 1
add_table(ws_sal, "جدول_المبيعات", f"A{SAL_HEAD_ROW}:I{SAL_TOTALS_ROW}")
tbl = ws_sal.tables["جدول_المبيعات"]
tbl.totalsRowShown = True
tbl.totalsRowCount = 1
tbl.autoFilter = AutoFilter(ref=f"A{SAL_HEAD_ROW}:I{SAL_LAST}")
for col in tbl.tableColumns:
    if col.name == "المبلغ":
        col.totalsRowFunction = "sum"
    if col.name == "البيان":
        col.totalsRowLabel = "الإجمالي الكلي"
plain_cell(ws_sal, f"E{SAL_TOTALS_ROW}", "الإجمالي الكلي", bold=True, align="right", fill=C_ACCENT)
formula_cell(ws_sal, f"F{SAL_TOTALS_ROW}", f"=SUBTOTAL(109,جدول_المبيعات[المبلغ])", FMT_MONEY, bold=True)
for col in "ABCDGHI":
    c = ws_sal[f"{col}{SAL_TOTALS_ROW}"]
    c.fill = PatternFill("solid", fgColor=C_ACCENT)
    c.border = BORDER_ALL

add_list_dv(ws_sal, f"G{SAL_FIRST}:G{SAL_LAST}", "=قائمة_طرق_الدفع")

# ---- ملخص أسبوعي ----
block_title(ws_sal, "K4:L4", "الملخص الأسبوعي")
header_cell(ws_sal, "K5", "الأسبوع"); header_cell(ws_sal, "L5", "الإجمالي")
formula_cell(ws_sal, "K6", '=IFERROR(SORT(UNIQUE(جدول_المبيعات[الأسبوع])),"")', align="right")
for i in range(60):
    rr = 6 + i
    formula_cell(ws_sal, f"L{rr}",
                 f'=IF($K{rr}="","",SUMIFS(جدول_المبيعات[المبلغ],جدول_المبيعات[الأسبوع],$K{rr}))',
                 FMT_MONEY)

# ---- ملخص شهري ----
block_title(ws_sal, "N4:O4", "الملخص الشهري")
header_cell(ws_sal, "N5", "الشهر"); header_cell(ws_sal, "O5", "الإجمالي")
formula_cell(ws_sal, "N6", '=IFERROR(SORT(UNIQUE(جدول_المبيعات[الشهر])),"")', align="right")
for i in range(30):
    rr = 6 + i
    formula_cell(ws_sal, f"O{rr}",
                 f'=IF($N{rr}="","",SUMIFS(جدول_المبيعات[المبلغ],جدول_المبيعات[الشهر],$N{rr}))',
                 FMT_MONEY)

# ---- بيانات مساعدة (كل العملاء) - مخفية ----
block_title(ws_sal, "Q4:R4", "بيانات مساعدة - كل العملاء")
header_cell(ws_sal, "Q5", "العميل"); header_cell(ws_sal, "R5", "الإجمالي")
formula_cell(ws_sal, "Q6", '=IFERROR(UNIQUE(جدول_المبيعات[العميل]),"")', align="right")
for i in range(60):
    rr = 6 + i
    formula_cell(ws_sal, f"R{rr}",
                 f'=IF($Q{rr}="","",SUMIFS(جدول_المبيعات[المبلغ],جدول_المبيعات[العميل],$Q{rr}))',
                 FMT_MONEY)
ws_sal.column_dimensions["Q"].hidden = True
ws_sal.column_dimensions["R"].hidden = True

# ---- أفضل 10 عملاء ----
block_title(ws_sal, "T4:V4", "أفضل 10 عملاء")
header_cell(ws_sal, "T5", "الترتيب"); header_cell(ws_sal, "U5", "العميل"); header_cell(ws_sal, "V5", "الإجمالي")
for i in range(10):
    rr = 6 + i
    plain_cell(ws_sal, f"T{rr}", i + 1, align="center")
    formula_cell(ws_sal, f"V{rr}", f'=IFERROR(LARGE($R$6:$R$65,T{rr}),"")', FMT_MONEY)
    formula_cell(ws_sal, f"U{rr}",
                 f'=IF(V{rr}="","",IFERROR(INDEX($Q$6:$Q$65,MATCH(V{rr},$R$6:$R$65,0)),""))',
                 align="right")

ws_sal.freeze_panes = f"A{SAL_FIRST}"
ws_sal.sheet_properties.tabColor = C_PRIMARY
print("sales sheet ok")

# ============================================================================
# ورقة المصروفات
# ============================================================================
ws_exp = new_sheet("المصروفات")
set_col_widths(ws_exp, {"A": 12, "B": 20, "C": 26, "D": 14, "E": 14, "F": 22,
                         "G": 10, "H": 3, "I": 22, "J": 13, "K": 14, "L": 3,
                         "M": 12, "N": 13})

big_title(ws_exp, "A1:G1", '="سجل المصروفات  -  " & اسم_الشركة', size=16)
sub_title(ws_exp, "A2:G2",
          '="الفترة الحالية: " & TEXT(الشهر_الحالي,"yyyy-mm") & "   |   العملة: " & العملة')

EXP_HEAD_ROW = 4
EXP_FIRST = 5
EXP_HEADERS = ["التاريخ", "الفئة", "البيان", "المبلغ", "طريقة الدفع", "ملاحظات", "الشهر"]
for i, h in enumerate(EXP_HEADERS):
    header_cell(ws_exp, f"{get_column_letter(1+i)}{EXP_HEAD_ROW}", h)

EXP_SAMPLE = [
    (datetime.date(2026, 5, 18), "رسوم حكومية", "رسوم رخصة المقاولات", 2500, "تحويل بنكي", ""),
    (datetime.date(2026, 6, 12), "إيجار", "إيجار المستودع شهر يونيو", 6000, "تحويل بنكي", ""),
    (datetime.date(2026, 6, 20), "صيانة المركبات", "تغيير زيت وفحص دوري", 900, "نقدي", ""),
    (datetime.date(2026, 7, 1), "ديزل", "تعبئة وقود المعدات", 1200, "نقدي", ""),
    (datetime.date(2026, 7, 3), "قطع غيار", "صيانة رافعة", 3400, "تحويل بنكي", ""),
    (datetime.date(2026, 7, 8), "اتصالات", "فواتير جوالات الموظفين", 450, "بطاقة", ""),
]
EXP_BLANK_ROWS = 15
EXP_LAST = EXP_FIRST + len(EXP_SAMPLE) + EXP_BLANK_ROWS - 1


def expenses_row_formula(ws, r):
    formula_cell(ws, f"G{r}", '=IF([@التاريخ]="","",TEXT([@التاريخ],"yyyy-mm"))')


r = EXP_FIRST
for d, cat, desc, amt, pm, note in EXP_SAMPLE:
    input_cell(ws_exp, f"A{r}", d, FMT_DATE)
    input_cell(ws_exp, f"B{r}", cat, align="right")
    input_cell(ws_exp, f"C{r}", desc, align="right")
    input_cell(ws_exp, f"D{r}", amt, FMT_MONEY)
    input_cell(ws_exp, f"E{r}", pm)
    input_cell(ws_exp, f"F{r}", note, align="right")
    expenses_row_formula(ws_exp, r)
    r += 1
for _ in range(EXP_BLANK_ROWS):
    input_cell(ws_exp, f"A{r}", None, FMT_DATE)
    input_cell(ws_exp, f"B{r}", None, align="right")
    input_cell(ws_exp, f"C{r}", None, align="right")
    input_cell(ws_exp, f"D{r}", None, FMT_MONEY)
    input_cell(ws_exp, f"E{r}")
    input_cell(ws_exp, f"F{r}", None, align="right")
    expenses_row_formula(ws_exp, r)
    r += 1

EXP_TOTALS_ROW = EXP_LAST + 1
add_table(ws_exp, "جدول_المصروفات", f"A{EXP_HEAD_ROW}:G{EXP_TOTALS_ROW}")
tbl = ws_exp.tables["جدول_المصروفات"]
tbl.totalsRowShown = True
tbl.totalsRowCount = 1
tbl.autoFilter = AutoFilter(ref=f"A{EXP_HEAD_ROW}:G{EXP_LAST}")
for col in tbl.tableColumns:
    if col.name == "المبلغ":
        col.totalsRowFunction = "sum"
    if col.name == "البيان":
        col.totalsRowLabel = "الإجمالي الكلي"
plain_cell(ws_exp, f"C{EXP_TOTALS_ROW}", "الإجمالي الكلي", bold=True, align="right", fill=C_ACCENT)
formula_cell(ws_exp, f"D{EXP_TOTALS_ROW}", f"=SUBTOTAL(109,جدول_المصروفات[المبلغ])", FMT_MONEY, bold=True)
for col in "ABEFG":
    c = ws_exp[f"{col}{EXP_TOTALS_ROW}"]
    c.fill = PatternFill("solid", fgColor=C_ACCENT)
    c.border = BORDER_ALL

add_list_dv(ws_exp, f"B{EXP_FIRST}:B{EXP_LAST}", "=قائمة_فئات_المصروفات")
add_list_dv(ws_exp, f"E{EXP_FIRST}:E{EXP_LAST}", "=قائمة_طرق_الدفع")

# ---- ملخص المصروفات حسب الفئة (يعتمد مباشرة على قائمة الفئات في الإعدادات) ----
block_title(ws_exp, "I4:K4", "ملخص المصروفات حسب الفئة")
header_cell(ws_exp, "I5", "الفئة"); header_cell(ws_exp, "J5", "الإجمالي"); header_cell(ws_exp, "K5", "النسبة")
formula_cell(ws_exp, "I6", "=الإعدادات!جدول_فئات_المصروفات[الفئة]", align="right")
EXP_CAT_SUMMARY_LAST = 6 + 24 - 1
for i in range(25):
    rr = 6 + i
    formula_cell(ws_exp, f"J{rr}",
                 f'=IF($I{rr}="","",SUMIFS(جدول_المصروفات[المبلغ],جدول_المصروفات[الفئة],$I{rr}))',
                 FMT_MONEY)
    formula_cell(ws_exp, f"K{rr}",
                 f'=IF($I{rr}="","",IFERROR(J{rr}/SUM(جدول_المصروفات[المبلغ]),0))', FMT_PCT)

# ---- الملخص الشهري ----
block_title(ws_exp, "M4:N4", "الملخص الشهري")
header_cell(ws_exp, "M5", "الشهر"); header_cell(ws_exp, "N5", "الإجمالي")
formula_cell(ws_exp, "M6", '=IFERROR(SORT(UNIQUE(جدول_المصروفات[الشهر])),"")', align="right")
for i in range(30):
    rr = 6 + i
    formula_cell(ws_exp, f"N{rr}",
                 f'=IF($M{rr}="","",SUMIFS(جدول_المصروفات[المبلغ],جدول_المصروفات[الشهر],$M{rr}))',
                 FMT_MONEY)

ws_exp.freeze_panes = f"A{EXP_FIRST}"
ws_exp.sheet_properties.tabColor = C_DANGER
print("expenses sheet ok")

# ============================================================================
# ورقة سجل عمل الموظفين
# ============================================================================
ws_log = new_sheet("سجل عمل الموظفين")
set_col_widths(ws_log, {"A": 12, "B": 16, "C": 15, "D": 15, "E": 12, "F": 13,
                         "G": 18, "H": 20, "I": 10, "J": 3,
                         "K": 16, "L": 20, "M": 20})

big_title(ws_log, "A1:I1", '="سجل عمل الموظفين اليومي  -  " & اسم_الشركة', size=16)
sub_title(ws_log, "A2:I2",
          '="ملاحظة: وقت الانصراف للتوثيق فقط ولا يدخل في أي حساب. يُحسب الإضافي فقط من (وقت انتهاء العمل) "'
          '&"بعد الساعة " & TEXT(وقت_نهاية_الدوام_الرسمي,"hh:mm") & '
          '" - وإذا انتهى العمل بعد منتصف الليل يُحتسب ضمن نفس يوم العمل تلقائياً."')

LOG_HEAD_ROW = 4
LOG_FIRST = 5
LOG_HEADERS = ["التاريخ", "اسم الموظف", "وقت الانصراف (للتوثيق فقط)", "وقت انتهاء العمل",
               "ساعات إضافية", "قيمة الإضافي", "اسم المشروع", "ملاحظات", "الشهر"]
for i, h in enumerate(LOG_HEADERS):
    header_cell(ws_log, f"{get_column_letter(1+i)}{LOG_HEAD_ROW}", h)

LOG_SAMPLE = [
    (datetime.date(2026, 7, 1), "إسماعيل", datetime.time(22, 30), datetime.time(23, 45), "مشروع فيلا الشمال", ""),
    (datetime.date(2026, 7, 1), "صادق", datetime.time(22, 0), datetime.time(23, 0), "مشروع فيلا الشمال", ""),
    (datetime.date(2026, 7, 2), "إسلام", datetime.time(23, 0), datetime.time(0, 30), "مشروع البرج التجاري", "استمر العمل بعد منتصف الليل"),
    (datetime.date(2026, 7, 3), "إسماعيل", datetime.time(21, 0), datetime.time(22, 30), "مشروع فيلا الشمال", ""),
    (datetime.date(2026, 7, 5), "صادق", datetime.time(22, 45), datetime.time(1, 15), "مشروع البرج التجاري", "استمر العمل بعد منتصف الليل"),
    (datetime.date(2026, 7, 6), "إسلام", datetime.time(22, 0), datetime.time(23, 30), "مشروع فيلا الشمال", ""),
]
LOG_BLANK_ROWS = 24
LOG_LAST = LOG_FIRST + len(LOG_SAMPLE) + LOG_BLANK_ROWS - 1

OT_HOURS_F = ('=IFERROR(IF([@[وقت انتهاء العمل]]=0,0,'
              'LET(e,[@[وقت انتهاء العمل]],off,وقت_نهاية_الدوام_الرسمي,thr,حد_اعتبار_الاستمرار,'
              'eff,IF(e<=thr,e+1,e),ROUND(MAX(0,(eff-off)*24),2))),0)')
OT_AMOUNT_F = ('=IFERROR([@[ساعات إضافية]]*XLOOKUP([@[اسم الموظف]],'
               'الإعدادات!جدول_الموظفين[اسم الموظف],الإعدادات!جدول_الموظفين[معدل الساعة الإضافية]),0)')


def log_row_formulas(ws, r):
    formula_cell(ws, f"E{r}", OT_HOURS_F, FMT_MONEY)
    formula_cell(ws, f"F{r}", OT_AMOUNT_F, FMT_MONEY)
    formula_cell(ws, f"I{r}", '=IF([@التاريخ]="","",TEXT([@التاريخ],"yyyy-mm"))')


r = LOG_FIRST
for d, emp, dep, end, proj, note in LOG_SAMPLE:
    input_cell(ws_log, f"A{r}", d, FMT_DATE)
    input_cell(ws_log, f"B{r}", emp, align="right")
    input_cell(ws_log, f"C{r}", dep, FMT_TIME)
    input_cell(ws_log, f"D{r}", end, FMT_TIME)
    input_cell(ws_log, f"G{r}", proj, align="right")
    input_cell(ws_log, f"H{r}", note, align="right")
    log_row_formulas(ws_log, r)
    r += 1
for _ in range(LOG_BLANK_ROWS):
    input_cell(ws_log, f"A{r}", None, FMT_DATE)
    input_cell(ws_log, f"B{r}", None, align="right")
    input_cell(ws_log, f"C{r}", None, FMT_TIME)
    input_cell(ws_log, f"D{r}", None, FMT_TIME)
    input_cell(ws_log, f"G{r}", None, align="right")
    input_cell(ws_log, f"H{r}", None, align="right")
    log_row_formulas(ws_log, r)
    r += 1

add_table(ws_log, "جدول_سجل_العمل", f"A{LOG_HEAD_ROW}:I{LOG_LAST}")
add_list_dv(ws_log, f"B{LOG_FIRST}:B{LOG_LAST}", "=قائمة_الموظفين_النشطين")

# ---- ملخص الإضافي الشهري لكل موظف (يُنقل تلقائياً إلى الرواتب) ----
block_title(ws_log, "K4:M4", "ملخص الإضافي للفترة الحالية (يُنقل تلقائياً لصفحة الرواتب)")
header_cell(ws_log, "K5", "الموظف"); header_cell(ws_log, "L5", "ساعات الإضافي"); header_cell(ws_log, "M5", "قيمة الإضافي")
formula_cell(ws_log, "K6", "=قائمة_الموظفين_النشطين", align="right")
for i in range(20):
    rr = 6 + i
    formula_cell(ws_log, f"L{rr}",
                 f'=IF($K{rr}="","",SUMIFS(جدول_سجل_العمل[ساعات إضافية],جدول_سجل_العمل[اسم الموظف],$K{rr},'
                 f'جدول_سجل_العمل[الشهر],TEXT(الشهر_الحالي,"yyyy-mm")))', FMT_MONEY)
    formula_cell(ws_log, f"M{rr}",
                 f'=IF($K{rr}="","",SUMIFS(جدول_سجل_العمل[قيمة الإضافي],جدول_سجل_العمل[اسم الموظف],$K{rr},'
                 f'جدول_سجل_العمل[الشهر],TEXT(الشهر_الحالي,"yyyy-mm")))', FMT_MONEY)

ws_log.freeze_panes = f"A{LOG_FIRST}"
ws_log.sheet_properties.tabColor = C_WARNING
print("work log sheet ok")

# ============================================================================
# ورقة الرواتب (الصفحة الرئيسية لإدارة الموظفين)
# ============================================================================
ws_pay = new_sheet("الرواتب")
set_col_widths(ws_pay, {"A": 16, "B": 13, "C": 13, "D": 13, "E": 13, "F": 12,
                         "G": 12, "H": 13, "I": 11, "J": 12, "K": 11, "L": 11,
                         "M": 13, "N": 13, "O": 13})

big_title(ws_pay, "A1:O1", '="كشف الرواتب الشهري  -  " & اسم_الشركة', size=16)
sub_title(ws_pay, "A2:O2",
          '="فترة الرواتب: " & TEXT(الشهر_الحالي,"yyyy-mm") & "   |   قائمة الموظفين تُبنى تلقائياً من ورقة الإعدادات"')

PAY_HEAD_ROW = 4
PAY_FIRST = 5
PAY_HEADERS = ["اسم الموظف", "الراتب الأساسي", "إجمالي ساعات الإضافي", "معدل الساعة الإضافية",
               "إجمالي قيمة الإضافي", "أيام العمل بالعطلة", "أجر يوم العطلة", "تعويض العطلات",
               "المكافآت", "بدلات أخرى", "السلف", "الخصومات", "إجمالي الراتب",
               "صافي الراتب", "حالة الدفع"]
for i, h in enumerate(PAY_HEADERS):
    header_cell(ws_pay, f"{get_column_letter(1+i)}{PAY_HEAD_ROW}", h)

PAY_ROWS = 15
PAY_LAST = PAY_FIRST + PAY_ROWS - 1

# عمود اسم الموظف: صيغة واحدة فقط في أول خلية تُنشر تلقائياً (Dynamic Array Spill)
formula_cell(ws_pay, f"A{PAY_FIRST}",
             '=IFERROR(قائمة_الموظفين_النشطين,"لا يوجد موظفون نشطون حالياً")', align="right", bold=True)

MONTH_TXT = 'TEXT(الشهر_الحالي,"yyyy-mm")'
for i in range(PAY_ROWS):
    r = PAY_FIRST + i
    if r != PAY_FIRST:
        # الخلايا تحت A{PAY_FIRST} تُملأ تلقائياً بالانسكاب - لا نكتب فيها شيئاً
        for col in "ABCDEFGHIJKLMNO":
            cell = ws_pay[f"{col}{r}"]
            cell.border = BORDER_ALL
            cell.font = Font(name=FONT_NAME, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center")

    formula_cell(ws_pay, f"B{r}",
                 f'=IF($A{r}="","",IFERROR(XLOOKUP($A{r},الإعدادات!جدول_الموظفين[اسم الموظف],'
                 f'الإعدادات!جدول_الموظفين[الراتب الأساسي]),0))', FMT_MONEY)
    formula_cell(ws_pay, f"C{r}",
                 f'=IF($A{r}="","",SUMIFS(جدول_سجل_العمل[ساعات إضافية],جدول_سجل_العمل[اسم الموظف],$A{r},'
                 f'جدول_سجل_العمل[الشهر],{MONTH_TXT}))', FMT_MONEY)
    formula_cell(ws_pay, f"D{r}",
                 f'=IF($A{r}="","",IFERROR(XLOOKUP($A{r},الإعدادات!جدول_الموظفين[اسم الموظف],'
                 f'الإعدادات!جدول_الموظفين[معدل الساعة الإضافية]),0))', FMT_MONEY)
    formula_cell(ws_pay, f"E{r}", f'=IF($A{r}="","",C{r}*D{r})', FMT_MONEY)
    input_cell(ws_pay, f"F{r}", None, FMT_INT)
    formula_cell(ws_pay, f"G{r}",
                 f'=IF($A{r}="","",IFERROR(XLOOKUP($A{r},الإعدادات!جدول_الموظفين[اسم الموظف],'
                 f'الإعدادات!جدول_الموظفين[أجر يوم العطلة]),0))', FMT_MONEY)
    formula_cell(ws_pay, f"H{r}", f'=IF($A{r}="","",F{r}*G{r})', FMT_MONEY)
    input_cell(ws_pay, f"I{r}", None, FMT_MONEY)
    input_cell(ws_pay, f"J{r}", None, FMT_MONEY)
    formula_cell(ws_pay, f"K{r}",
                 f'=IF($A{r}="","",SUMIFS(جدول_حركات_الموظفين[المبلغ],جدول_حركات_الموظفين[الموظف],$A{r},'
                 f'جدول_حركات_الموظفين[نوع الحركة],"سلفة",جدول_حركات_الموظفين[الشهر],{MONTH_TXT}))', FMT_MONEY)
    formula_cell(ws_pay, f"L{r}",
                 f'=IF($A{r}="","",SUMIFS(جدول_حركات_الموظفين[المبلغ],جدول_حركات_الموظفين[الموظف],$A{r},'
                 f'جدول_حركات_الموظفين[نوع الحركة],"خصم",جدول_حركات_الموظفين[الشهر],{MONTH_TXT}))', FMT_MONEY)
    formula_cell(ws_pay, f"M{r}", f'=IF($A{r}="","",B{r}+E{r}+H{r}+I{r}+J{r})', FMT_MONEY, bold=True)
    formula_cell(ws_pay, f"N{r}", f'=IF($A{r}="","",M{r}-K{r}-L{r})', FMT_MONEY, bold=True,
                 color=C_SUCCESS)
    input_cell(ws_pay, f"O{r}", None)

add_list_dv(ws_pay, f"O{PAY_FIRST}:O{PAY_LAST}", '"مدفوع,غير مدفوع,جزئي"')

# صف الإجمالي
TOT_ROW = PAY_LAST + 2
plain_cell(ws_pay, f"A{TOT_ROW}", "الإجمالي", bold=True, align="right", fill=C_ACCENT)
SUM_COLS = "BCEHIJKLMN"
for col in SUM_COLS:
    formula_cell(ws_pay, f"{col}{TOT_ROW}", f'=SUM({col}{PAY_FIRST}:{col}{PAY_LAST})', FMT_MONEY, bold=True)
for col in "ABCDEFGHIJKLMNO":
    if col not in SUM_COLS and col != "A":
        ws_pay[f"{col}{TOT_ROW}"].fill = PatternFill("solid", fgColor=C_ACCENT)
        ws_pay[f"{col}{TOT_ROW}"].border = BORDER_ALL

# تنسيق شرطي لحالة الدفع
ws_pay.conditional_formatting.add(
    f"O{PAY_FIRST}:O{PAY_LAST}",
    FormulaRule(formula=[f'O{PAY_FIRST}="مدفوع"'], fill=PatternFill("solid", fgColor=C_SUCCESS_L)))
ws_pay.conditional_formatting.add(
    f"O{PAY_FIRST}:O{PAY_LAST}",
    FormulaRule(formula=[f'O{PAY_FIRST}="غير مدفوع"'], fill=PatternFill("solid", fgColor=C_DANGER_L)))
ws_pay.conditional_formatting.add(
    f"O{PAY_FIRST}:O{PAY_LAST}",
    FormulaRule(formula=[f'O{PAY_FIRST}="جزئي"'], fill=PatternFill("solid", fgColor=C_WARNING_L)))

# ============================================================================
# جدول حركات الموظفين (سلف / خصومات) - أسفل صفحة الرواتب
# ============================================================================
TX_TITLE_ROW = TOT_ROW + 3
TX_HEAD_ROW = TX_TITLE_ROW + 1
TX_FIRST = TX_HEAD_ROW + 1

big_title(ws_pay, f"A{TX_TITLE_ROW}:G{TX_TITLE_ROW}", "حركات السلف والخصومات", size=13, fill=C_PRIMARY2, height=24)

TX_HEADERS = ["التاريخ", "الموظف", "نوع الحركة", "المبلغ", "السبب", "ملاحظات", "الشهر"]
for i, h in enumerate(TX_HEADERS):
    header_cell(ws_pay, f"{get_column_letter(1+i)}{TX_HEAD_ROW}", h)

TX_SAMPLE = [
    (datetime.date(2026, 7, 2), "إسماعيل", "سلفة", 500, "ظرف طارئ", ""),
    (datetime.date(2026, 7, 10), "صادق", "خصم", 100, "تأخير عن العمل", ""),
]
TX_BLANK_ROWS = 13
TX_LAST = TX_FIRST + len(TX_SAMPLE) + TX_BLANK_ROWS - 1

r = TX_FIRST
for d, emp, ttype, amt, reason, note in TX_SAMPLE:
    input_cell(ws_pay, f"A{r}", d, FMT_DATE)
    input_cell(ws_pay, f"B{r}", emp, align="right")
    input_cell(ws_pay, f"C{r}", ttype)
    input_cell(ws_pay, f"D{r}", amt, FMT_MONEY)
    input_cell(ws_pay, f"E{r}", reason, align="right")
    input_cell(ws_pay, f"F{r}", note, align="right")
    formula_cell(ws_pay, f"G{r}", '=IF([@التاريخ]="","",TEXT([@التاريخ],"yyyy-mm"))')
    r += 1
for _ in range(TX_BLANK_ROWS):
    input_cell(ws_pay, f"A{r}", None, FMT_DATE)
    input_cell(ws_pay, f"B{r}", None, align="right")
    input_cell(ws_pay, f"C{r}")
    input_cell(ws_pay, f"D{r}", None, FMT_MONEY)
    input_cell(ws_pay, f"E{r}", None, align="right")
    input_cell(ws_pay, f"F{r}", None, align="right")
    formula_cell(ws_pay, f"G{r}", '=IF([@التاريخ]="","",TEXT([@التاريخ],"yyyy-mm"))')
    r += 1

add_table(ws_pay, "جدول_حركات_الموظفين", f"A{TX_HEAD_ROW}:G{TX_LAST}")
add_list_dv(ws_pay, f"B{TX_FIRST}:B{TX_LAST}", "=قائمة_الموظفين_النشطين")
add_list_dv(ws_pay, f"C{TX_FIRST}:C{TX_LAST}", '"سلفة,خصم"')

ws_pay.conditional_formatting.add(
    f"C{TX_FIRST}:C{TX_LAST}",
    FormulaRule(formula=[f'C{TX_FIRST}="سلفة"'], fill=PatternFill("solid", fgColor=C_WARNING_L)))
ws_pay.conditional_formatting.add(
    f"C{TX_FIRST}:C{TX_LAST}",
    FormulaRule(formula=[f'C{TX_FIRST}="خصم"'], fill=PatternFill("solid", fgColor=C_DANGER_L)))

ws_pay.freeze_panes = f"A{PAY_FIRST}"
ws_pay.sheet_properties.tabColor = C_SUCCESS
print("payroll sheet ok")

# ============================================================================
# لوحة التحكم (Dashboard) - أول ورقة في المصنف
# ============================================================================
ws_dash = new_sheet("لوحة التحكم")
set_col_widths(ws_dash, {"A": 3})
for col in "BCDEFGHIJKLMNOP":
    ws_dash.column_dimensions[col].width = 11
ws_dash.column_dimensions["A"].width = 3

big_title(ws_dash, "A1:P1", '="لوحة التحكم الرئيسية  -  " & اسم_الشركة', size=20, height=36)
sub_title(ws_dash, "A2:P2",
          '="فترة التقرير: " & TEXT(الشهر_الحالي,"yyyy-mm") & "   |   السنة المالية: " & السنة_المالية &'
          ' "   |   العملة: " & العملة', size=11)
ws_dash.row_dimensions[3].height = 6


def kpi_card(top_row, col_start, col_end, label, value_formula, fmt, fill, value_color=C_TEXT,
             suffix=""):
    lc = get_column_letter(col_start)
    rc = get_column_letter(col_end)
    label_range = f"{lc}{top_row}:{rc}{top_row}"
    value_range = f"{lc}{top_row+1}:{rc}{top_row+3}"
    ws_dash.merge_cells(label_range)
    lab = ws_dash[f"{lc}{top_row}"]
    lab.value = label
    lab.font = Font(name=FONT_NAME, size=10, bold=True, color=C_WHITE)
    lab.fill = PatternFill("solid", fgColor=fill)
    lab.alignment = Alignment(horizontal="center", vertical="center")
    for cc in ws_dash[label_range][0]:
        cc.fill = PatternFill("solid", fgColor=fill)
    ws_dash.row_dimensions[top_row].height = 20

    ws_dash.merge_cells(value_range)
    val = ws_dash[f"{lc}{top_row+1}"]
    val.value = xf(value_formula) if value_formula.startswith("=") else value_formula
    val.font = Font(name=FONT_NAME, size=16, bold=True, color=value_color)
    val.fill = PatternFill("solid", fgColor=C_WHITE)
    val.alignment = Alignment(horizontal="center", vertical="center")
    val.number_format = fmt
    for row_cells in ws_dash[value_range]:
        for cc in row_cells:
            cc.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cc.fill = PatternFill("solid", fgColor=C_WHITE)
    for rr in range(top_row + 1, top_row + 4):
        ws_dash.row_dimensions[rr].height = 18


MONTH_TXT_D = 'TEXT(الشهر_الحالي,"yyyy-mm")'

# صف البطاقات الأول
kpi_card(5, 2, 4, "المبيعات الشهرية",
         f'=SUMIFS(المبيعات!جدول_المبيعات[المبلغ],المبيعات!جدول_المبيعات[الشهر],{MONTH_TXT_D})',
         FMT_MONEY, C_PRIMARY, C_SUCCESS)
kpi_card(5, 6, 8, "المصروفات الشهرية",
         f'=SUMIFS(المصروفات!جدول_المصروفات[المبلغ],المصروفات!جدول_المصروفات[الشهر],{MONTH_TXT_D})',
         FMT_MONEY, C_PRIMARY, C_DANGER)
kpi_card(5, 10, 12, "صافي الربح", "=B6-F6", FMT_MONEY, C_PRIMARY, C_TEXT)
kpi_card(5, 14, 16, "صافي رواتب الموظفين", f"=الرواتب!N{TOT_ROW}", FMT_MONEY, C_PRIMARY, C_TEXT)

# صف البطاقات الثاني
kpi_card(9, 2, 4, "إجمالي قيمة الإضافي", f"=الرواتب!E{TOT_ROW}", FMT_MONEY, C_PRIMARY2, C_TEXT)
kpi_card(9, 6, 8, "إجمالي السلف", f"=الرواتب!K{TOT_ROW}", FMT_MONEY, C_PRIMARY2, C_WARNING)
kpi_card(9, 10, 12, "إجمالي الخصومات", f"=الرواتب!L{TOT_ROW}", FMT_MONEY, C_PRIMARY2, C_DANGER)
kpi_card(9, 14, 16, "عدد الموظفين النشطين",
         '=COUNTIFS(الإعدادات!جدول_الموظفين[الحالة],"نشط")', FMT_INT, C_PRIMARY2, C_TEXT)

# بطاقات نصية عريضة
kpi_card(13, 2, 8, "أعلى فئة مصروفات",
         '=IFERROR(INDEX(المصروفات!$I$6:$I$30,MATCH(MAX(المصروفات!$J$6:$J$30),المصروفات!$J$6:$J$30,0)),'
         '"لا توجد بيانات")', "@", C_WARNING, C_TEXT)
kpi_card(13, 10, 16, "أفضل عميل",
         '=IFERROR(المبيعات!U6,"لا توجد بيانات")', "@", C_WARNING, C_TEXT)

# تنسيق شرطي لبطاقة صافي الربح (أخضر إن كان موجباً، أحمر إن كان سالباً)
ws_dash.conditional_formatting.add(
    "J6:L8", FormulaRule(formula=["$J$6<0"], font=Font(name=FONT_NAME, size=16, bold=True, color=C_DANGER)))
ws_dash.conditional_formatting.add(
    "J6:L8", FormulaRule(formula=["$J$6>=0"], font=Font(name=FONT_NAME, size=16, bold=True, color=C_SUCCESS)))

ws_dash.row_dimensions[17].height = 8

# ---- الصفوف الفارغة الفاصلة قبل الرسوم البيانية ----
block_title(ws_dash, "B18:P18", "الاتجاهات الشهرية (آخر 12 شهراً) والتوزيع حسب الفئة")

# ---- بيانات مساعدة للرسوم البيانية (أعمدة R:U) ----
header_cell(ws_dash, "R4", "الشهر")
header_cell(ws_dash, "S4", "التسمية")
header_cell(ws_dash, "T4", "المبيعات")
header_cell(ws_dash, "U4", "المصروفات")
header_cell(ws_dash, "V4", "الربح")
TREND_MONTHS = 12
for i in range(TREND_MONTHS):
    rr = 5 + i
    formula_cell(ws_dash, f"R{rr}", f'=EDATE(الشهر_الحالي,{i - (TREND_MONTHS - 1)})', "yyyy-mm")
    formula_cell(ws_dash, f"S{rr}",
                 f'=CHOOSE(MONTH(R{rr}),"يناير","فبراير","مارس","أبريل","مايو","يونيو",'
                 f'"يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر")&" "&YEAR(R{rr})', align="right")
    formula_cell(ws_dash, f"T{rr}",
                 f'=SUMIFS(المبيعات!جدول_المبيعات[المبلغ],المبيعات!جدول_المبيعات[الشهر],TEXT(R{rr},"yyyy-mm"))',
                 FMT_MONEY)
    formula_cell(ws_dash, f"U{rr}",
                 f'=SUMIFS(المصروفات!جدول_المصروفات[المبلغ],المصروفات!جدول_المصروفات[الشهر],TEXT(R{rr},"yyyy-mm"))',
                 FMT_MONEY)
    formula_cell(ws_dash, f"V{rr}", f"=T{rr}-U{rr}", FMT_MONEY)
TREND_LAST = 5 + TREND_MONTHS - 1
for col in "RSTUV":
    ws_dash.column_dimensions[col].hidden = True

print("dashboard base ok")

# ---- الرسم البياني الأول: اتجاه المبيعات والمصروفات والربح (أعمدة + خط) ----
bar = BarChart()
bar.type = "col"
bar.style = 10
bar.title = "اتجاه المبيعات والمصروفات (12 شهراً)"
bar.y_axis.title = None
bar.x_axis.title = None
bar.height = 9
bar.width = 22
bar.visible_cells_only = False

cats = Reference(ws_dash, min_col=19, min_row=5, max_row=TREND_LAST)          # S column
sales_ref = Reference(ws_dash, min_col=20, min_row=4, max_row=TREND_LAST)     # T column incl header
exp_ref = Reference(ws_dash, min_col=21, min_row=4, max_row=TREND_LAST)       # U column incl header
bar.add_data(sales_ref, titles_from_data=True)
bar.add_data(exp_ref, titles_from_data=True)
bar.set_categories(cats)
bar.series[0].graphicalProperties.solidFill = C_PRIMARY2
bar.series[1].graphicalProperties.solidFill = C_DANGER

line = LineChart()
profit_ref = Reference(ws_dash, min_col=22, min_row=4, max_row=TREND_LAST)   # V column incl header
line.add_data(profit_ref, titles_from_data=True)
line.set_categories(cats)
line.series[0].graphicalProperties.line.width = 25000
line.series[0].graphicalProperties.line.solidFill = C_SUCCESS
line.series[0].smooth = False
line.visible_cells_only = False

bar += line
ws_dash.add_chart(bar, "B19")

# ---- الرسم البياني الثاني: توزيع المصروفات حسب الفئة ----
pie = PieChart()
pie.title = "توزيع المصروفات حسب الفئة"
pie.height = 9
pie.width = 14
pie.visible_cells_only = False
cat_labels = Reference(ws_exp, min_col=9, min_row=6, max_row=25)
cat_values = Reference(ws_exp, min_col=10, min_row=5, max_row=25)
pie.add_data(cat_values, titles_from_data=True)
pie.set_categories(cat_labels)
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
ws_dash.add_chart(pie, "K19")

ws_dash.sheet_view.showGridLines = False
ws_dash.sheet_properties.tabColor = C_PRIMARY
print("dashboard charts ok")

# ============================================================================
# دليل الاستخدام
# ============================================================================
ws_help = new_sheet("دليل الاستخدام")
set_col_widths(ws_help, {"A": 3, "B": 34, "C": 78})

big_title(ws_help, "A1:C1", "دليل استخدام نظام إدارة شركة المقاولات", size=18, height=34)
sub_title(ws_help, "A2:C2", "نظام إدارة متكامل - مصمم ليكون سهل الاستخدام دون الحاجة لخبرة في إكسل", size=11)
ws_help.row_dimensions[3].height = 8


def guide_section(row, title):
    block_title(ws_help, f"B{row}:C{row}", title, fill=C_PRIMARY2, size=12)
    return row + 1


def guide_line(row, label, text):
    plain_cell(ws_help, f"B{row}", label, bold=True, align="right", color=C_PRIMARY)
    plain_cell(ws_help, f"C{row}", text, align="right")
    ws_help[f"B{row}"].alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
    ws_help[f"C{row}"].alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
    ws_help.row_dimensions[row].height = 30
    return row + 1


row = 5
row = guide_section(row, "1. فكرة عمل النظام")
row = guide_line(row, "التصميم", "كل صفحة من صفحات النظام (المبيعات، المصروفات، سجل العمل، الرواتب، لوحة التحكم) "
                                 "تسحب بياناتها تلقائياً من ورقة إعدادات واحدة مخفية اسمها (الإعدادات). "
                                 "لا حاجة لإدخال اسم موظف أو فئة مصروف أو راتب في أكثر من مكان.")
row = guide_line(row, "الشهر الحالي", "توجد خلية واحدة في الإعدادات تتحكم بكل أرقام الشهر في المصنف بالكامل "
                                      "(المبيعات الشهرية، المصروفات الشهرية، الرواتب، لوحة التحكم). "
                                      "لإعداد رواتب شهر جديد، غيّر هذه الخلية فقط وستتحدث كل الصفحات تلقائياً.")
row += 1

row = guide_section(row, "2. كيفية إضافة أو تعديل موظف")
row = guide_line(row, "الخطوات", "1) انقر بزر الفأرة الأيمن على أي اسم ورقة أسفل الشاشة ثم اختر (إظهار). "
                                 "2) اختر ورقة (الإعدادات). 3) أضف اسم الموظف وراتبه ومعدلاته في جدول الموظفين. "
                                 "سيظهر الموظف تلقائياً في القوائم المنسدلة وفي صفحة الرواتب دون أي تعديل إضافي.")
row = guide_line(row, "إيقاف موظف", "غيّر عمود (الحالة) الخاص به إلى (غير نشط) وسيختفي تلقائياً من صفحة الرواتب "
                                    "ومن قوائم الاختيار في سجل العمل وحركات السلف والخصومات.")
row += 1

row = guide_section(row, "3. الألوان ودلالتها")
row = guide_line(row, "خلايا صفراء", "خلايا إدخال يدوي - هذه هي الخلايا الوحيدة التي يجب الكتابة فيها.")
row = guide_line(row, "خلايا رمادية", "خلايا محسوبة تلقائياً بمعادلات ومحمية من التعديل - لا تحتاج لأي إدخال.")
row = guide_line(row, "حالة الدفع", "أخضر = مدفوع  /  أحمر = غير مدفوع  /  أصفر = جزئي.")
row += 1

row = guide_section(row, "4. سجل عمل الموظفين والإضافي")
row = guide_line(row, "وقت الانصراف", "هذا الحقل للتوثيق فقط ولا يدخل في أي عملية حسابية.")
row = guide_line(row, "احتساب الإضافي", 'يُحسب فقط من "وقت انتهاء العمل" بعد وقت انتهاء الدوام الرسمي (23:00 افتراضياً). '
                                        "إذا انتهى العمل بعد منتصف الليل (مثال 12:30 ص) يتعامل النظام معه تلقائياً "
                                        "كاستمرار لنفس يوم العمل وليس يوماً جديداً.")
row = guide_line(row, "الترحيل للرواتب", "إجمالي ساعات وقيمة الإضافي لكل موظف عن الشهر الحالي يُنقل تلقائياً "
                                         "إلى صفحة الرواتب دون أي إدخال يدوي.")
row += 1

row = guide_section(row, "5. السلف والخصومات")
row = guide_line(row, "طريقة الإدخال", "تُسجَّل كل سلفة أو خصم كصف واحد في جدول (حركات السلف والخصومات) أسفل صفحة الرواتب. "
                                       "يقوم النظام تلقائياً بجمعها لكل موظف وخصمها ضمن صافي الراتب - لا تُدخل يدوياً في جدول الرواتب.")
row += 1

row = guide_section(row, "6. الحماية")
row = guide_line(row, "خلايا محمية", "أوراق العمل محمية بكلمة مرور لمنع التعديل غير المقصود على المعادلات، "
                                     "مع ترك خلايا الإدخال (الصفراء) مفتوحة دائماً للكتابة والفرز والتصفية.")
row = guide_line(row, "كلمة المرور", PROTECT_PASSWORD)
row += 1

row = guide_section(row, "7. قائمة أوراق العمل")
row = guide_line(row, "لوحة التحكم", "ملخص تنفيذي بأهم الأرقام والرسوم البيانية المحدثة تلقائياً.")
row = guide_line(row, "المبيعات", "سجل الفواتير والمبيعات مع ملخص أسبوعي وشهري وأفضل 10 عملاء.")
row = guide_line(row, "المصروفات", "سجل المصروفات مع ملخص حسب الفئة والنسبة المئوية والملخص الشهري.")
row = guide_line(row, "سجل عمل الموظفين", "تسجيل الحضور اليومي واحتساب الإضافي تلقائياً.")
row = guide_line(row, "الرواتب", "كشف الرواتب الشهري الكامل + جدول السلف والخصومات.")

ws_help.sheet_properties.tabColor = C_GRAY_TEXT
ws_help.sheet_view.showGridLines = False
print("guide sheet ok")

# ============================================================================
# ترتيب الأوراق + الحماية + الحفظ النهائي
# ============================================================================
ORDER = ["لوحة التحكم", "المبيعات", "المصروفات", "سجل عمل الموظفين",
         "الرواتب", "دليل الاستخدام", "الإعدادات"]
wb._sheets = [wb[name] for name in ORDER]
wb.active = 0

PROTECTED_SHEETS = ["لوحة التحكم", "المبيعات", "المصروفات", "سجل عمل الموظفين", "الرواتب"]
for name in PROTECTED_SHEETS:
    ws = wb[name]
    ws.protection.sheet = True
    ws.protection.password = PROTECT_PASSWORD
    ws.protection.autoFilter = False
    ws.protection.sort = False
    ws.protection.formatCells = False
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False

import os
OUT_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                         "نظام_إدارة_شركة_المقاولات.xlsx")
wb.save(OUT_PATH)
print("FINAL SAVE OK ->", OUT_PATH)
