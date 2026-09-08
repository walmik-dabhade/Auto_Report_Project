# funder_report_generator_v5.py
# 14-page CSR / Community Education Programme report generator
# Reads the cleaned Excel workbook and translated_report.json.
# Run from the project root:
#   python report\funder_report_generator_v5.py

import json
import math
from pathlib import Path
from collections import OrderedDict

import pandas as pd
import matplotlib.pyplot as plt

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


ROOT = Path(__file__).resolve().parents[1]
# This generator is designed for the original 2025-26 workbook.
# Expected local project layout:
#   Auto_Report_Project\data\MONTH DATA FILE-2025-26 - complete(1).xlsx
# A fallback search is included so a small filename difference does not break the run.
INPUT_CANDIDATES = [
    ROOT / "data" / "MONTH DATA FILE-2025-26 - complete(1).xlsx",
    ROOT / "data" / "MONTH DATA FILE-2025-26 - complete.xlsx",
    ROOT / "MONTH DATA FILE-2025-26 - complete(1).xlsx",
    ROOT / "MONTH DATA FILE-2025-26 - complete.xlsx",
]
INPUT_XLSX = next((p for p in INPUT_CANDIDATES if p.exists()), None)
if INPUT_XLSX is None:
    matches = list(ROOT.rglob("MONTH DATA FILE-2025-26*.xlsx"))
    if matches:
        INPUT_XLSX = matches[0]



TRANSLATED_JSON = ROOT / "output" / "translated_report.json"
if not TRANSLATED_JSON.exists():
    matches = list(ROOT.rglob("translated_report.json"))
    if matches:
        TRANSLATED_JSON = matches[0]
OUTPUT_DIR = ROOT / "output" / "funder_reports_v5"
CHART_DIR = ROOT / "output" / "v5_charts"
IMAGE_ROOT = ROOT / "images"
ASSET_ROOT = ROOT / "assets"
LOGO_ROOTS = [ROOT / "assets" / "logos", ROOT / "report" / "assets" / "logos", ROOT / "images" / "logos"]
DSSF_LOGO = next((r / "DSSF.png" for r in LOGO_ROOTS if (r / "DSSF.png").exists()), None)


FUNDERS = ["YARDI", "NICE", "BREMBO", "ACS", "IDRF", "BBD",
           "GHATKOPAR", "DSSF", "AVAYA"]

ACTIVITY_ORDER = [
    "BALWADI", "STUDY CLASS", "LIBRARY CLASS(S)",
    "HOME LANDING", "LIBRARY CLASS"
]


def clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def pct(n, d):
    return round((n / d) * 100, 2) if d else 0.0


def norm_level(v):
    s = clean(v).upper().replace(" ", "")
    if s in {"L1", "L1.1", "L1.2"}: return "L1"
    if s in {"L2", "L2.1", "L2.2"}: return "L2"
    if s in {"L3", "L3.1", "L3.2"}: return "L3"
    if s in {"L4", "L4.1", "L4.2", "L4.3", "L4.4", "L4.5", "L4.6"}: return "L4"
    return ""


def norm_math(v):
    s = clean(v).upper().replace(" ", "")
    for i in range(1, 10):
        if s.startswith(f"M{i}"):
            return f"M{i}"
    return ""


def status(v):
    s = clean(v).upper()
    if "CONTINU" in s: return "Continued"
    if "DROUP" in s or "DROP" in s: return "Dropout"
    if "MIGRAT" in s: return "Migrated"
    if "GRADUAT" in s: return "Graduate"
    return ""


def load_data():
    """Load the original workbook and standardise only fields used by the report."""
    if INPUT_XLSX is None or not INPUT_XLSX.exists():
        raise FileNotFoundError("Original workbook not found. Put MONTH DATA FILE-2025-26 - complete(1).xlsx in the project data folder.")

    xls = pd.ExcelFile(INPUT_XLSX)
    preferred_sheets = ["CHILDRENS NAME", "CHILDREN NAME", "CLEANED_CHILDREN_DATA", "CHILDREN NAME"]
    sheet = next((name for name in preferred_sheets if name in xls.sheet_names), None)
    if sheet is None:
        # Find a sheet that actually contains the December status field.
        for name in xls.sheet_names:
            try:
                cols = pd.read_excel(INPUT_XLSX, sheet_name=name, nrows=0).columns
                if "STATUS DECEMBER-25" in cols:
                    sheet = name
                    break
            except Exception:
                continue
    if sheet is None:
        raise ValueError("No worksheet containing December programme data was found.")
    df = pd.read_excel(INPUT_XLSX, sheet_name=sheet, header=0)

    # Sanity check: the full report dataset must contain the December fields.
    if "STATUS DECEMBER-25" not in df.columns:
        raise ValueError(
            "The selected workbook sheet is not the full April–December dataset. "
            "The report requires the 'CHILDRENS NAME' sheet containing December columns."
        )

    required = [
        "FUNDING WISE", "UPDATED CLASS TYPE", "GENDER", "SOFTWEAR ID",
        "STATUS DECEMBER-25", "LANGUAGE LEVEL DECEMBER-25",
        "MATHS LEVEL DECEMBER-25", "LL BASELINE", "ML BASELINE",
        "CATEGORY.8", "CATEGORY.10", "STANDARD"
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("The workbook is missing required columns: " + ", ".join(missing))

    df["FUNDING WISE"] = df["FUNDING WISE"].map(clean).str.upper()
    df = df[df["FUNDING WISE"].isin(FUNDERS)].copy()
    df["GENDER"] = df["GENDER"].map(clean).str.title()
    df["UPDATED CLASS TYPE"] = (
        df["UPDATED CLASS TYPE"].map(clean).str.upper()
        .str.replace(r"\s+", " ", regex=True)
    )
    df["SOFTWEAR ID"] = df["SOFTWEAR ID"].map(clean)
    df["DEC_STATUS"] = df["STATUS DECEMBER-25"].map(status)
    df["DEC_LANG"] = df["LANGUAGE LEVEL DECEMBER-25"].map(norm_level)
    df["DEC_MATH"] = df["MATHS LEVEL DECEMBER-25"].map(norm_math)
    df["BASE_LANG"] = df["LL BASELINE"].map(norm_level)
    df["BASE_MATH"] = df["ML BASELINE"].map(norm_math)
    df["ATT_CAT"] = df["CATEGORY.8"].map(clean).str.upper()
    df["LEARNING_DAYS"] = df["CATEGORY.10"].map(clean).str.upper()
    df["STANDARD_CLEAN"] = df["STANDARD"].map(clean).str.upper()
    return df


def load_translation():
    if not TRANSLATED_JSON.exists():
        return {}
    try:
        return json.loads(TRANSLATED_JSON.read_text(encoding="utf-8"))
    except Exception:
        return {}


def choose_story(data, preferred=None):
    stories = data.get("success_stories", [])
    if isinstance(stories, dict):
        stories = list(stories.values())
    if not isinstance(stories, list):
        return ""
    for s in stories:
        if isinstance(s, str):
            return s
        if isinstance(s, dict):
            title = clean(s.get("title") or s.get("name") or "")
            body = clean(s.get("story") or s.get("description") or
                         s.get("details") or s.get("narrative") or "")
            if preferred and preferred.lower() in title.lower():
                return body or title
            if body:
                return body
    return ""


def get_section_text(data, *keys):
    cur = data
    for k in keys:
        if isinstance(cur, dict):
            cur = cur.get(k, "")
        else:
            return ""
    if isinstance(cur, str):
        return cur
    if isinstance(cur, list):
        return " ".join(clean(x) for x in cur if clean(x))
    if isinstance(cur, dict):
        vals = []
        for v in cur.values():
            if isinstance(v, str):
                vals.append(v)
        return " ".join(vals)
    return ""


def funder_logo_path(funder):
    for root in LOGO_ROOTS:
        for ext in (".png", ".jpg", ".jpeg"):
            p = root / f"{funder}{ext}"
            if p.exists():
                return p
    return None


def set_page_border(section):
    """Black double-line page border matching Report Sample 2."""
    sectPr = section._sectPr
    pgBorders = sectPr.first_child_found_in("w:pgBorders")
    if pgBorders is None:
        pgBorders = OxmlElement("w:pgBorders")
        sectPr.append(pgBorders)
    pgBorders.set(qn("w:offsetFrom"), "page")
    for edge in ("top", "left", "bottom", "right"):
        tag = qn(f"w:{edge}")
        el = pgBorders.find(tag)
        if el is None:
            el = OxmlElement(f"w:{edge}")
            pgBorders.append(el)
        el.set(qn("w:val"), "double")
        el.set(qn("w:sz"), "18")
        el.set(qn("w:space"), "8")
        el.set(qn("w:color"), "000000")


def add_sample_header_footer(section):
    header = section.header
    # Clear default paragraph and use a two-column header table.
    header.paragraphs[0].text = ""
    tab = header.add_table(rows=1, cols=2, width=Inches(6.9))
    tab.alignment = WD_TABLE_ALIGNMENT.CENTER
    for cell in tab.rows[0].cells:
        cell.width = Inches(3.45)
    p = tab.cell(0,0).paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r=p.add_run("October 2025 – December 2025"); r.font.name="Arial"; r.font.size=Pt(10)
    p = tab.cell(0,1).paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r=p.add_run("Community Education Program"); r.font.name="Arial"; r.font.size=Pt(10)
    footer = section.footer
    footer.paragraphs[0].text = ""
    ft = footer.add_table(rows=1, cols=2, width=Inches(6.9))
    ft.alignment = WD_TABLE_ALIGNMENT.CENTER
    p=ft.cell(0,0).paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    r=p.add_run("Door Step School Foundation"); r.font.name="Arial"; r.font.size=Pt(10)
    p=ft.cell(0,1).paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    r=p.add_run("Page "); r.font.name="Arial"; r.font.size=Pt(10)
    fld=OxmlElement("w:fldSimple"); fld.set(qn("w:instr"),"PAGE"); p._p.append(fld)


def add_logo_or_text(cell, path, text, width):
    p=cell.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(10)
    if path and path.exists():
        p.add_run().add_picture(str(path), width=Inches(width))
    else:
        r=p.add_run(text); r.bold=True; r.font.name="Arial"; r.font.size=Pt(20)


def add_cover_layout(doc, donor, funder):
    logo_table=doc.add_table(rows=1, cols=2); logo_table.alignment=WD_TABLE_ALIGNMENT.CENTER; logo_table.autofit=False
    logo_table.columns[0].width=Inches(3.35); logo_table.columns[1].width=Inches(3.35)
    add_logo_or_text(logo_table.cell(0,0), DSSF_LOGO, "DOOR STEP SCHOOL\nFOUNDATION", 1.55)
    add_logo_or_text(logo_table.cell(0,1), funder_logo_path(funder), donor, 1.75)

    band=doc.add_table(rows=1, cols=1); band.alignment=WD_TABLE_ALIGNMENT.CENTER
    c=band.cell(0,0); shade(c,"5B9BD5")
    p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(10)
    r=p.add_run(f"Community Education Programme supported by\n{donor}"); r.bold=True; r.font.name="Arial"; r.font.size=Pt(18); r.font.color.rgb=RGBColor(255,255,255)

    band2=doc.add_table(rows=1, cols=1); band2.alignment=WD_TABLE_ALIGNMENT.CENTER
    c=band2.cell(0,0); shade(c,"E2F0D9")
    p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(10)
    r=p.add_run("Project Progress Report"); r.bold=True; r.font.name="Arial"; r.font.size=Pt(23)
    p=c.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(10)
    r=p.add_run("October 2025 – December 2025"); r.bold=True; r.font.name="Arial"; r.font.size=Pt(17)

    imgs=[]
    if IMAGE_ROOT.exists():
        imgs=sorted(p for p in IMAGE_ROOT.rglob("*") if p.suffix.lower() in {".jpg",".jpeg",".png"})
    if imgs:
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(18); p.paragraph_format.space_after=Pt(10)
        p.add_run().add_picture(str(imgs[0]), width=Inches(5.85))
    else:
        placeholder(doc,"Cover programme photograph",2.85)

    box=doc.add_table(rows=1, cols=1); box.alignment=WD_TABLE_ALIGNMENT.RIGHT; box.autofit=False; box.columns[0].width=Inches(3.05)
    c=box.cell(0,0); set_cell_text(c,"Door Step School Foundation\n110, Parimal, Anand Park, Aundh,\nPune 411007.\nPhone: 91-9823859002\nE mail: info@dssf.org.in\nWebsite: www.dssf.org.in",True,10)
    for p in c.paragraphs: p.alignment=WD_ALIGN_PARAGRAPH.RIGHT

def add_page_number(section):
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run("Page ")
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    footer._p.append(fld)


def shade(cell, fill="D9EAF7"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def set_cell_text(cell, text, bold=False, size=10):
    size = max(float(size), 10.0)
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(str(text))
    r.bold = bold
    r.font.name = "Arial"
    r.font.size = Pt(size)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in("w:tcMar")
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for side in ("top", "start", "bottom", "end"):
        node = tcMar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tcMar.append(node)
        node.set(qn("w:w"), "55")
        node.set(qn("w:type"), "dxa")


def add_table(doc, headers, rows, widths=None, font=10.0):
    font = max(float(font), 10.0)
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.autofit = False
    hdr = table.rows[0]
    hdr._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
    hdr._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    for i, h in enumerate(headers):
        set_cell_text(hdr.cells[i], h, True, font)
        shade(hdr.cells[i], "5B9BD5")
        for run in hdr.cells[i].paragraphs[0].runs:
            run.font.color.rgb = RGBColor(255,255,255)
    for ri, row in enumerate(rows):
        tr = table.add_row()
        tr._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        for i, v in enumerate(row):
            set_cell_text(tr.cells[i], v, False, font)
            shade(tr.cells[i], "DDEBF7" if ri % 2 == 0 else "FFFFFF")
    if rows and str(rows[-1][0]).strip().lower() == "total":
        for c in table.rows[-1].cells:
            shade(c, "4472C4")
            for run in c.paragraphs[0].runs:
                run.bold=True; run.font.color.rgb=RGBColor(255,255,255)
    if widths:
        for row in table.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Inches(w)
    return table


def add_title(doc, text, size=15):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(max(float(size), 10.0))
    return p


def add_subtitle(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(10.5)
    return p


def add_body(doc, text, size=10.0):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.05
    r = p.add_run(text)
    r.font.size = Pt(size)
    return p


def add_bullets(doc, items, size=10.0):
    for item in items:
        p = doc.add_paragraph(style=None)
        p.paragraph_format.left_indent = Inches(0.18)
        p.paragraph_format.first_line_indent = Inches(-0.12)
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.line_spacing = 1.0
        r = p.add_run("• " + item)
        r.font.size = Pt(size)


def page_break(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.add_run().add_break(WD_BREAK.PAGE)
    return p



def placeholder(doc, label, height=1.25):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_text(cell, f"[ PHOTO SPACE ]\n{label}", True, 10)
    cell.height = Inches(height)
    shade(cell, "F2F2F2")
    return table


def available_images():
    imgs=[]
    for root in [IMAGE_ROOT, ASSET_ROOT]:
        if root.exists():
            imgs.extend([p for p in root.rglob("*") if p.suffix.lower() in {".jpg",".jpeg",".png"} and "logo" not in str(p).lower()])
    return sorted(set(imgs))


def photo_grid(doc, labels, height=2.35):
    imgs=available_images()
    if imgs:
        cols=2
        table=doc.add_table(rows=math.ceil(len(labels)/cols), cols=cols)
        table.alignment=WD_TABLE_ALIGNMENT.CENTER
        idx=0
        for r in range(len(table.rows)):
            for c in range(cols):
                cell=table.cell(r,c); cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
                if idx < len(imgs):
                    p=cell.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_after=Pt(2)
                    p.add_run().add_picture(str(imgs[idx]),width=Inches(3.05))
                    idx+=1
                else:
                    set_cell_text(cell,"",False,10)
        return table
    table=doc.add_table(rows=math.ceil(len(labels)/2), cols=2)
    table.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,label in enumerate(labels):
        cell=table.cell(i//2,i%2); set_cell_text(cell,f"[ PHOTO SPACE ]\n{label}",True,10)
        cell.height=Inches(height); shade(cell,"F2F2F2")
        cell.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    return table



def make_chart_funder(df, funder, outdir):
    outdir.mkdir(parents=True, exist_ok=True)
    paths = {}

    # Gender — sample-style pie chart
    paths["gender"] = add_pie_chart_funder(df, funder, outdir)

    # Attendance
    cats = ["80 TO 100", "51 TO 79", "1 TO 50"]
    vals = [int((df["ATT_CAT"] == c).sum()) for c in cats]
    fig, ax = plt.subplots(figsize=(5.5, 1.9))
    ax.bar(["High", "Medium", "Low"], vals)
    ax.set_title("Attendance distribution")
    ax.set_ylabel("Children")
    fig.tight_layout()
    p = outdir / f"{funder}_attendance.png"
    fig.savefig(p, dpi=180)
    plt.close(fig)
    paths["attendance"] = p

    # Baseline vs December language
    levels = ["L1", "L2", "L3", "L4"]
    base = [int((df["BASE_LANG"] == x).sum()) for x in levels]
    dec = [int((df["DEC_LANG"] == x).sum()) for x in levels]
    fig, ax = plt.subplots(figsize=(5.7, 2.15))
    x = range(len(levels))
    width = 0.36
    ax.bar([i-width/2 for i in x], base, width, label="Baseline")
    ax.bar([i+width/2 for i in x], dec, width, label="December")
    ax.set_xticks(list(x), levels)
    ax.set_ylabel("Children")
    ax.set_title("Language learning levels: baseline vs December")
    ax.legend()
    fig.tight_layout()
    p = outdir / f"{funder}_language_comparison.png"
    fig.savefig(p, dpi=180)
    plt.close(fig)
    paths["language"] = p

    # Activity reach
    activity_names = []
    activity_vals = []
    for a in ACTIVITY_ORDER:
        n = int((df["UPDATED CLASS TYPE"] == a).sum())
        if n:
            activity_names.append(a.title())
            activity_vals.append(n)
    if activity_names:
        fig, ax = plt.subplots(figsize=(6.0, 2.5))
        ax.bar(activity_names, activity_vals)
        ax.set_ylabel("Children")
        ax.set_title("Activity-wise reach")
        ax.tick_params(axis="x", rotation=25)
        fig.tight_layout()
        p = outdir / f"{funder}_activity.png"
        fig.savefig(p, dpi=180)
        plt.close(fig)
        paths["activity"] = p

    return paths


def metrics(df):
    total = len(df)
    cont = int((df.DEC_STATUS == "Continued").sum())
    drop = int((df.DEC_STATUS == "Dropout").sum())
    mig = int((df.DEC_STATUS == "Migrated").sum())
    grad = int((df.DEC_STATUS == "Graduate").sum())
    return {
        "total": total, "continued": cont, "dropout": drop,
        "migrated": mig, "graduate": grad,
        "cont_pct": pct(cont, total), "drop_pct": pct(drop, total),
        "mig_pct": pct(mig, total)
    }


def activity_rows(df):
    rows = []
    for a in ACTIVITY_ORDER:
        x = df[df["UPDATED CLASS TYPE"] == a]
        if x.empty:
            continue
        m = metrics(x)
        rows.append([a.title(), m["total"], m["continued"], m["dropout"],
                     m["migrated"], m["graduate"]])
    return rows


def daywise_language_rows(df):
    """Language levels by learning-day band."""
    raw = df["LEARNING_DAYS"].map(clean).str.upper()
    aliases = OrderedDict([
        ("1 to 50", {"1 TO 50", "1-50", "1 TO 50 DAYS"}),
        ("51 to 90", {"51 TO 90", "51-90", "51 TO 90 DAYS"}),
        ("91 to 120", {"91 TO 120", "91-120", "91 TO 120 DAYS"}),
        ("Above 120", {"ABOVE 120", "ABOVE 120 DAYS", "121 TO 150", "121-150"}),
    ])
    rows = []
    for label, values in aliases.items():
        x = df[raw.isin(values)]
        rows.append([
            label,
            int((x["DEC_LANG"] == "L1").sum()),
            int((x["DEC_LANG"] == "L2").sum()),
            int((x["DEC_LANG"] == "L3").sum()),
            int((x["DEC_LANG"] == "L4").sum()),
            len(x),
        ])
    rows.append([
        "Total",
        int((df["DEC_LANG"] == "L1").sum()),
        int((df["DEC_LANG"] == "L2").sum()),
        int((df["DEC_LANG"] == "L3").sum()),
        int((df["DEC_LANG"] == "L4").sum()),
        len(df),
    ])
    return rows


def safe_source_text(translated, *keys, limit=1200):
    text = get_section_text(translated, *keys)
    return text[:limit] if text else ""


def funder_display_name(funder):
    return {
        "YARDI": "Yardi Software",
        "NICE": "NICE",
        "BREMBO": "Brembo",
        "ACS": "ACS",
        "IDRF": "IDRF",
        "BBD": "BBD",
        "GHATKOPAR": "Ghatkopar",
        "DSSF": "DSSF",
        "AVAYA": "Avaya",
    }.get(funder, funder)


def add_cover_title(doc, text, size, space_after=10):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.bold = True
    r.font.name = "Arial"
    r.font.size = Pt(size)
    return p


def add_highlight_box(doc, text, yellow_phrase=None):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    shade(cell, "DDEBF7")
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(10)
    p.paragraph_format.left_indent = Inches(0.08)
    p.add_run("✓  ").bold = True
    if yellow_phrase and yellow_phrase in text:
        before, after = text.split(yellow_phrase, 1)
        p.add_run(before)
        r = p.add_run(yellow_phrase)
        r.bold = True
        rpr = r._r.get_or_add_rPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), "FFF2CC")
        rpr.append(shd)
        if after:
            p.add_run(after)
    else:
        r = p.add_run(text)
        r.bold = True
    return table


def add_section_heading_icon(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run("✚   ")
    r.bold = True
    r.font.size = Pt(11)
    r2 = p.add_run(text)
    r2.bold = True
    r2.font.size = Pt(14)
    return p


def add_underlined_heading(doc, text, size=10.0):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.bold = True
    r.underline = True
    r.font.size = Pt(max(float(size), 10.0))
    return p


def add_pie_chart_funder(df, funder, outdir):
    outdir.mkdir(parents=True, exist_ok=True)
    g = df["GENDER"].value_counts()
    vals = [int(g.get("Male", 0)), int(g.get("Female", 0))]
    labels = [
        f"Boys, {vals[0]}, {pct(vals[0], len(df)):.0f}%",
        f"Girls, {vals[1]}, {pct(vals[1], len(df)):.0f}%"
    ]
    fig, ax = plt.subplots(figsize=(3.2, 3.0))
    wedges, _ = ax.pie(
        vals, startangle=90, counterclock=False,
        wedgeprops={"edgecolor": "white", "linewidth": 1.2}
    )
    for w, label in zip(wedges, labels):
        ang = (w.theta2 + w.theta1) / 2
        x = 0.57 * math.cos(math.radians(ang))
        y = 0.57 * math.sin(math.radians(ang))
        ax.text(x, y, label, ha="center", va="center",
                fontsize=10.0, fontweight="bold", color="white")
    ax.set_aspect("equal")
    fig.tight_layout()
    path = outdir / f"{funder}_gender_pie.png"
    fig.savefig(path, dpi=220, bbox_inches="tight")
    plt.close(fig)
    return path


def grade_1_to_4_summary(df):
    vals = df["STANDARD_CLEAN"]
    n = int(vals.isin({"1", "2", "3", "4"}).sum())
    return n, pct(n, len(df))


def study_attendance_summary(df):
    x = df[df["UPDATED CLASS TYPE"] == "STUDY CLASS"]
    total = len(x)
    high = int((x["ATT_CAT"] == "80 TO 100").sum())
    return high, total, pct(high, total)


def qualitative_continuity_note(translated, funder):
    candidates = [
        get_section_text(translated, "challenges"),
        get_section_text(translated, "activity_stories", "study_class"),
        get_section_text(translated, "activity_stories", "library_class"),
        get_section_text(translated, "activity_stories", "home_lending"),
    ]
    for c in candidates:
        if c:
            return c[:900]
    return ("Regular follow-up with children and parents supported continuity, "
            "while attendance, school schedules and family movement remained "
            "important factors affecting participation.")

def add_activity_block(doc, title, objective, activities, reach, observations):
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    c=t.cell(0,0); shade(c,"5B9BD5"); set_cell_text(c,title,True,10)
    for run in c.paragraphs[0].runs: run.font.color.rgb=RGBColor(255,255,255)
    c=t.add_row().cells[0]; set_cell_text(c,"Objective –\n"+objective,False,10)
    c=t.add_row().cells[0]; set_cell_text(c,"Activities –",True,10)
    c=t.add_row().cells[0]; c.text=""
    p=c.paragraphs[0]; p.paragraph_format.space_after=Pt(2); p.paragraph_format.line_spacing=1.0
    for i,item in enumerate(activities,1):
        r=p.add_run(f"{i}. {item}" + ("\n" if i<len(activities) else "")); r.font.name="Arial"; r.font.size=Pt(10)
    c=t.add_row().cells[0]; set_cell_text(c,"Reach –\n"+reach,False,10)
    c=t.add_row().cells[0]; set_cell_text(c,"Observations –\n"+observations,False,10)
    for row in t.rows: row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    return t


def build_report(funder, all_df, translated):
    df = all_df[all_df["FUNDING WISE"] == funder].copy()
    m = metrics(df)
    charts = make_chart_funder(df, funder, CHART_DIR)
    donor = funder_display_name(funder)

    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.55)
    sec.bottom_margin = Inches(0.55)
    sec.left_margin = Inches(0.55)
    sec.right_margin = Inches(0.55)
    sec.different_first_page_header_footer = True
    set_page_border(sec)
    add_sample_header_footer(sec)

    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10.5)

    # PAGE 1 — Report Sample 2 style cover
    add_cover_layout(doc, donor, funder)
    page_break(doc)

    # PAGE 2 — EXACT SAMPLE WORDING; ONLY FUNDER + SUPPORTED COUNT CHANGE
    add_section_heading_icon(doc, "Introduction")
    add_body(doc,
        "Door Step School Foundation aims to address illiteracy among children (3 to 14 years) from "
        "marginalised sections of society like those from construction sites, pavement dwellers, temporary "
        "and permanent slums. Our focus is to address three major problems of public education system "
        "through various innovative programs:", 10.0)
    add_bullets(doc, [
        "Non-enrolment i.e. children are not enrolled in schools for a variety of reasons.",
        "Children drop out of school at an early age.",
        "Even if children are enrolled in schools, their quality of education is low and children suffer from low learning levels"
    ], 10.0)
    add_body(doc,
        "DSSF has also expanded its scope to take vocational counselling and guidance to school children "
        "between 12 to 14 years and youth in communities up to 25 years.", 10.0)

    add_section_heading_icon(doc, "Executive Summary")
    add_body(doc,
        "The education of children from migrant communities is frequently neglected, as they face challenges "
        "due to frequent migration, and the unfavourable conditions of their living in urban outskirts. Thus, "
        "they either remain non-enrolled, or become school drop-outs if enrolled. Their lower learning levels "
        "further hamper their educational continuity. More than often, they are forced to work alongside their "
        "parents, beg, and/or take care of household and younger siblings. It results in disconnect and lack of "
        "interest in formal schooling. Even though some of them are street-smart, these children often fall into "
        "the same cycle of poverty, are susceptible to addictions, and at risk of falling victims to anti-social elements.", 10.0)
    add_body(doc,
        "To generate interest and encourage their ongoing learning, it is necessary to adopt an innovative and "
        "consistent approach and use engaging teaching methods. It is crucial to focus on building their "
        "foundational literacy and numeracy skills to ensure their retention in school. DSS launched the "
        "Community Education Program which aims to take Foundational Literacy and Numeracy (FLN) skills to "
        "children which is in line with the FLN mission - NEP 2020. The program focuses on improving learning "
        "levels, overall development of children along with increasing parents’ participation towards continued "
        "education and hence a brighter future.", 10.0)
    add_body(doc,
        "The UN Sustainable Development Goal No. 4 focuses on ensuring inclusive and equitable quality education "
        "and promoting lifelong learning opportunities for all. Universal literacy and numeracy for all youth and "
        "substantial adults is one of the targets of SDG 4 – Quality education. The Community Education Program "
        "works towards achieving universal foundational literacy and numeracy in line with SDG -4 and as "
        "prioritized by NEP 2020 and Nipun Bharat Mission.", 10.0)
    add_body(doc,
        f"We thank {donor} for supporting the education of {m['total']} children between October and December "
        "2025 through Community Education Program in Lakshmi Nagar community located in Kondhwa area of Pune, Maharashtra.", 10.0)
    page_break(doc)

    # PAGE 3 — SAMPLE-MATCHED LAYOUT
    add_section_heading_icon(doc, "Highlights and Reach:")

    add_highlight_box(
        doc,
        f"We reached out to {m['total']} children from the supported programme area through different activities "
        "conducted between October 2025 and December 2025.",
        f"{m['total']} children"
    )

    assessed = int(df["DEC_LANG"].isin(["L1", "L2", "L3", "L4"]).sum())
    expected = int(df["DEC_LANG"].isin(["L3", "L4"]).sum())
    math_m7 = int((df["DEC_MATH"] == "M7").sum())
    add_highlight_box(
        doc,
        f"{pct(expected, assessed):.0f}% ({expected}/{assessed}) children are at L3/L4 among children with a "
        f"recorded December language level. {assessed} children have a recorded language level while "
        f"{math_m7} children are at M7 in numeracy.",
        f"{pct(expected, assessed):.0f}%"
    )

    add_underlined_heading(doc, "Activity-wise overall reach:")

    outer = doc.add_table(rows=1, cols=2)
    outer.alignment = WD_TABLE_ALIGNMENT.CENTER
    outer.autofit = False
    outer.columns[0].width = Inches(4.55)
    outer.columns[1].width = Inches(2.55)
    left = outer.cell(0, 0)
    right = outer.cell(0, 1)

    # Educational table.
    t = left.add_table(rows=1, cols=2)
    t.style = "Table Grid"
    set_cell_text(t.rows[0].cells[0], "Activity", True, 10.0)
    set_cell_text(t.rows[0].cells[1], "No. of Children", True, 10.0)
    shade(t.rows[0].cells[0], "5B9BD5")
    shade(t.rows[0].cells[1], "5B9BD5")
    educational = [
        ("Balwadi", "BALWADI"),
        ("Study Class", "STUDY CLASS"),
        ("Library (S)", "LIBRARY CLASS(S)"),
        ("Home Lending Books", "HOME LANDING"),
    ]
    for label, key in educational:
        n = int((df["UPDATED CLASS TYPE"] == key).sum())
        cells = t.add_row().cells
        set_cell_text(cells[0], label, False, 10.0)
        set_cell_text(cells[1], n, True, 10.0)
        shade(cells[0], "DDEBF7")
        shade(cells[1], "DDEBF7")
    cells = t.add_row().cells
    set_cell_text(cells[0], "Total", True, 10.0)
    set_cell_text(cells[1], m["total"], True, 10.0)
    for c in cells:
        shade(c, "4472C4")
        for run in c.paragraphs[0].runs:
            run.font.color.rgb = __import__("docx").shared.RGBColor(255, 255, 255)

    # Sample support rows are retained; do not fabricate funder-specific counts.
    support = ["Project (Balwadi)", "Science Experiments", "Events", "Exposure Visits", "Project Supanth"]
    support_values = {}
    if isinstance(translated, dict):
        for key in ["support_activities", "activity_participation", "support"]:
            obj = translated.get(key)
            if isinstance(obj, dict):
                support_values.update(obj)
    for label in support:
        value = support_values.get(label) or support_values.get(label.lower()) or "—"
        cells = t.add_row().cells
        set_cell_text(cells[0], label, False, 10.0)
        set_cell_text(cells[1], value, False, 10.0)
        shade(cells[0], "E2F0D9")
        shade(cells[1], "E2F0D9")

    p = right.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(charts["gender"]), width=Inches(2.25))
    p = right.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("*Average participants per session")
    r.italic = True
    r.font.size = Pt(10)

    add_underlined_heading(doc, "Children’s Continuity Status:")
    cont_rows = []
    for label, key in [
        ("Balwadi", "BALWADI"),
        ("Study Class", "STUDY CLASS"),
        ("Library (S)", "LIBRARY CLASS(S)"),
        ("Home Lending\nBooks", "HOME LANDING"),
    ]:
        x = df[df["UPDATED CLASS TYPE"] == key]
        cont_rows.append([
            label, len(x),
            int((x["DEC_STATUS"] == "Continued").sum()),
            int((x["DEC_STATUS"] == "Dropout").sum()),
            int((x["DEC_STATUS"] == "Migrated").sum()),
        ])
    cont_rows.append(["Total", m["total"], m["continued"], m["dropout"], m["migrated"]])
    add_table(doc, ["Activity", "No. of Children", "Continued", "Drop-out", "Migrated"], cont_rows, font=10.0)

    g14, gpct = grade_1_to_4_summary(df)
    high, study_total, highpct = study_attendance_summary(df)
    note = qualitative_continuity_note(translated, funder)
    add_bullets(doc, [
        f"{g14} ({gpct:.0f}%) children are studying in 1st to 4th grades.",
        f"{m['continued']}/{m['total']} ({m['cont_pct']:.0f}%) children are currently continuing classes. "
        f"{high}/{study_total} ({highpct:.0f}%) Study Class children have 80% & above attendance.",
        note
    ], 10.0)
    page_break(doc)

    # PAGE 4 — sample-matched activity objectives and observations
    add_section_heading_icon(doc, "Activity-wise Objectives, Reach and Observations")
    add_body(doc, "Different activities conducted by DSSF team under Community Education Program include –", 10)

    bal=df[df["UPDATED CLASS TYPE"]=="BALWADI"]
    add_activity_block(doc, "Balwadi (Pre-school)",
        "Prepare 3–6-year-old children from the communities for formal schooling.",
        [
            "Conduct engaging activities that support physical, social, emotional, intellectual, language and early numeracy development, involving parents and siblings whenever possible.",
            "Track each child’s progress and motivate parents to enrol children at the right age."
        ],
        f"{len(bal)} children are part of Balwadi activity." if len(bal) else "No Balwadi children are recorded for this funder.",
        safe_source_text(translated,"activity_stories","balwadi",limit=650) or "Children learn through engaging activities that support holistic development, early language and numeracy, confidence and school readiness.")

    study=df[df["UPDATED CLASS TYPE"]=="STUDY CLASS"]
    add_activity_block(doc, "Study Classes – Language (Marathi) Skills and Numeracy Skills for 6–14-year-old children",
        "Develop children’s foundational language and numeracy skills to make sure they learn basics and do not fall behind in studies and drop out of school.",
        [
            "Conduct study classes in line with the 150-day DSS pedagogy using interesting and engaging TLM and teaching methods.",
            "Assess children’s learning progress by tracking their learning levels and group children according to their current learning level.",
            "Strengthen language and numeracy skills so children can keep up with school studies and reduce the chances of dropping out."
        ],
        f"{len(study)} children are part of Study Class for this funder." if len(study) else "No Study Class children are recorded for this funder.",
        safe_source_text(translated,"activity_stories","study_class",limit=650) or "Teachers use level-based grouping, daily and monthly planning, charts, worksheets, books and activity-based methods to support children at their current learning level.")

    add_subtitle(doc,"Class Plan")
    add_table(doc,["Component","Implementation"],[
        ["Batch size","20–30 children per batch"],
        ["Session","Approximately 1 hour language + 1 hour numeracy"],
        ["Pedagogy","150-day structured approach; level-based grouping; daily and monthly planning"],
        ["Resources","Charts, worksheets, books, reading and numeracy corners"],
    ],font=10)
    page_break(doc)

    # PAGE 5 — keep language heading and table together
    add_section_heading_icon(doc,"Language Skills and Learning Progress")
    add_subtitle(doc,"Language learning-level distribution by learning days")
    add_table(doc,["Learning days","L1","L2","L3","L4","Total"],daywise_language_rows(df),font=10)
    assessed_f=int(df["DEC_LANG"].isin(["L1","L2","L3","L4"]).sum())
    l12=int(df["DEC_LANG"].isin(["L1","L2"]).sum())
    l34=int(df["DEC_LANG"].isin(["L3","L4"]).sum())
    add_bullets(doc,[
        f"{int((df['DEC_LANG']=='L4').sum())} children have reached L4 (Reading Practice).",
        f"{l34}/{assessed_f} ({pct(l34,assessed_f):.0f}%) assessed children are at L3–L4 learning levels.",
        f"{l12}/{assessed_f} ({pct(l12,assessed_f):.0f}%) assessed children remain at L1–L2 and require continued foundational support."
    ],10)
    add_subtitle(doc,"Language Skills: observations")
    add_body(doc,safe_source_text(translated,"learning_outcomes","language",limit=750) or "Children are grouped by current learning level. Teachers use repeated practice, graded materials and home reinforcement to help children move towards confident reading.",10)
    add_subtitle(doc,"Attendance of Children")
    high,study_total,highpct=study_attendance_summary(df)
    add_body(doc,(f"{m['continued']}/{m['total']} ({m['cont_pct']:.0f}%) children are currently continuing classes. {high}/{study_total} ({highpct:.0f}%) Study Class children have 80% & above attendance." if study_total else f"{m['continued']}/{m['total']} ({m['cont_pct']:.0f}%) children are currently continuing classes."),10)
    att_rows=[]
    for label,key in [("High (80–100%)","80 TO 100"),("Medium (51–79%)","51 TO 79"),("Low (1–50%)","1 TO 50")]:
        n=int((df["ATT_CAT"]==key).sum()); att_rows.append([label,n,f"{pct(n,len(df)):.0f}%"])
    att_rows.append(["Total",len(df),"100%" if len(df) else "0%"])
    add_table(doc,["Attendance category","No. of Children","Share"],att_rows,font=10)
    add_subtitle(doc,"To improve children’s attendance")
    add_bullets(doc,[
        "We created WhatsApp groups for parents and appreciated children and parents who attended regularly.",
        "We appreciated children with 100% attendance and used simple recognition to motivate others.",
        "Attendance charts helped children mark their own presence and notice their progress.",
        "Parent meetings and home visits were used to discuss progress, attendance and practical barriers.",
    ],10)
    page_break(doc)

    # PAGE 6 — comparative graph, numeracy and learning improvement
    add_section_heading_icon(doc,"Learning Outcomes and Numeracy Skills")
    add_subtitle(doc,"The graph below shows comparison of children’s baseline and current learning levels:")
    if charts.get("language"):
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.add_run().add_picture(str(charts["language"]),width=Inches(5.15))
    add_subtitle(doc,"Numeracy Skills")
    add_body(doc,"We use games, easy-to-solve numeric puzzles and easily available materials while teaching numeracy concepts.",10)
    mc=df["DEC_MATH"].value_counts(); math_assessed=int(df["DEC_MATH"].isin(["M1","M2","M3","M4","M5","M6","M7"]).sum())
    add_table(doc,["Learning Levels →","M1","M2","M3","M4","M5","M6","M7","Total"],[["No. of Children",*[int(mc.get(x,0)) for x in ["M1","M2","M3","M4","M5","M6","M7"]],math_assessed]],font=10)
    m7=int(mc.get("M7",0)); add_body(doc,f"{m7} ({pct(m7,math_assessed)}%) children are currently at M7.",10)
    add_subtitle(doc,"To improve children’s learning levels–")
    add_bullets(doc,[
        "We created new teaching tools while using existing DSS teaching-learning materials.",
        "We told children how many days it usually takes to complete each level, so they started tracking progress and setting goals.",
        "We gave worksheets for home practice and shared revision activities through WhatsApp groups.",
        "We shared DSS level-based books at home to help children practise."
    ],10)
    add_subtitle(doc,"Home Lending Books Activity")
    add_body(doc,safe_source_text(translated,"activity_stories","home_lending",limit=900) or "Books were shared regularly to build reading habits, strengthen vocabulary and connect reading with everyday learning.",10)
    page_break(doc)

    add_title(doc, "Attendance, Library/Home Lending and Stakeholder Engagement")
    add_subtitle(doc, "Attendance improvement")
    add_bullets(doc, [
        "Recognise regular attendance and encourage children through positive appreciation.",
        "Use attendance charts and simple rewards to make regular participation visible.",
        "Follow up with parents when attendance falls and understand practical barriers.",
        "Encourage regular children to bring other children from the neighbourhood."
    ], 10.0)
    add_subtitle(doc, "Library / Home Lending highlights")
    hl = get_section_text(translated, "activity_stories", "home_lending") or get_section_text(translated, "activity_stories", "home_landing")
    add_body(doc, hl[:1100] if hl else "Books were made available through library and home-lending activities to encourage reading beyond class time.")
    add_subtitle(doc, "Stakeholder engagement with children")
    add_body(doc, "Children participated in science experiments, events and projects designed to encourage curiosity, observation, explanation and independent participation.")

    add_subtitle(doc, "Science experiments")
    add_body(doc, get_section_text(translated, "activity_stories", "study_class")[:700] or "Hands-on experiments were used to make concepts concrete and encourage children to observe, discuss and explain.")
    add_subtitle(doc, "Events")
    add_body(doc, "Festivals, national days and child-centred events provided opportunities for participation, creativity, leadership and expression.")
    add_subtitle(doc, "Projects")
    add_body(doc, "Projects were used to connect classroom learning with children’s surroundings and encourage curiosity and independent work.")
    page_break(doc)

    add_title(doc, "Parent Engagement and DSSF Team Capacity Building")
    add_subtitle(doc, "Parent meetings and home visits")
    add_body(doc, get_section_text(translated, "parent_engagement")[:950] or "Parent interactions focused on regular attendance, children’s learning progress, home support and communication with teachers.")
    add_subtitle(doc, "Changes observed")
    add_bullets(doc, [
        "Parents increasingly discussed children’s studies and attendance with programme staff.",
        "Home visits supported follow-up with children who were irregular or absent.",
        "Parents were encouraged to support reading and learning practice at home.",
        "Regular communication helped teachers understand family-level barriers."
    ], 10.0)
    add_subtitle(doc, "DSSF Team Capacity Building")
    add_body(doc, get_section_text(translated, "teacher_training")[:850] or "Training sessions strengthened teachers’ classroom practice, record keeping, learning-level support, parent engagement and activity planning.")

    add_title(doc, "Challenges, Plan for Next Quarter and Special Story")
    add_subtitle(doc, "Challenges")
    add_bullets(doc, [
        "Changes in school timings can affect continuity and attendance.",
        "School holidays can reduce regular participation.",
        "Some parents may not send children regularly even after repeated interaction and counselling.",
        "Migration and changes in family circumstances can affect continuity."
    ], 10)
    add_subtitle(doc, "Plan for next quarter")
    add_bullets(doc, [
        "Work with children on improving learning levels and reading fluency.",
        "Strengthen regular class attendance through parent follow-up and positive reinforcement.",
        "Prioritise children at lower learning levels for additional support.",
        "Improve completeness of learning and attendance records."
    ], 10)
    add_subtitle(doc, "Special Story")
    add_body(doc, choose_story(translated)[:850] or "A learner-facing success story from the translated field report should be inserted here.", 10)
    page_break(doc)

    add_title(doc, "Annexures – Highlights")
    add_subtitle(doc, "Annexure 1: Community Profile")
    community = get_section_text(translated, "programme_overview", "community_profile")
    add_body(doc, community[:1400] or "Community profile information should be drawn from the translated field report and retained as qualitative context.")
    add_subtitle(doc, "Annexure 2: Learning Level Descriptions")
    add_table(doc, ["Language level", "Description"], [
        ["L1", "Learning alphabets"], ["L2", "Learning matras"],
        ["L3", "Learning composite letters"], ["L4", "Reading practice"]
    ], font=10)
    add_subtitle(doc, "Annexure 3: Activities with children")
    add_bullets(doc, ["Science experiments", "Events and celebrations", "Projects", "Library and home lending", "Study Class learning activities"], 10)
    page_break(doc)

    add_title(doc, "Annexure 4: Parent Meeting Topics")
    parent = translated.get("parent_engagement", {})
    month_rows = []
    if isinstance(parent, dict):
        monthly = parent.get("monthly") or parent.get("month_wise") or parent.get("details")
        if isinstance(monthly, list):
            for r in monthly:
                if isinstance(r, dict):
                    month_rows.append([
                        clean(r.get("month")), clean(r.get("topic") or r.get("topics")),
                        clean(r.get("parents_attended") or r.get("parents")),
                        clean(r.get("home_visits"))
                    ])
    if not month_rows:
        month_rows = [
            ["January", "Sending children regularly to class", "Source report", "Source report"],
            ["February", "DSS teaching methodology and learning progress", "Source report", "Source report"],
            ["March", "TLM exhibition and use of teaching-learning materials", "Source report", "Source report"],
            ["April", "Regular attendance of children", "Source report", "Source report"],
            ["May", "Summer break", "—", "—"],
            ["June", "Introducing new parents to DSSF work", "Source report", "Source report"],
            ["July", "Attendance and learning levels of children", "Source report", "Source report"],
            ["August", "Attendance and learning levels of children", "Source report", "Source report"],
            ["September", "Exam preparation", "Source report", "Source report"],
            ["October", "Diwali vacation and regular attendance", "Source report", "Source report"],
            ["November", "Smartphone usage – pros and cons", "Source report", "Source report"],
            ["December", "Class attendance and learning levels", "Source report", "Source report"],
        ]
    add_table(doc, ["Month", "Parent meeting topics", "No. of parents attended", "No. of home visits"], month_rows, font=10.0)
    page_break(doc)

    add_title(doc, "Annexure 5: DSSF Team Training Topics")
    training = translated.get("teacher_training", {})
    month_rows = []
    if isinstance(training, dict):
        monthly = training.get("monthly") or training.get("month_wise") or training.get("details")
        if isinstance(monthly, list):
            for r in monthly:
                if isinstance(r, dict):
                    month_rows.append([
                        clean(r.get("month")), clean(r.get("topic") or r.get("training_topic")),
                        clean(r.get("participants") or r.get("no_of_participants"))
                    ])
    if not month_rows:
        month_rows = [
            ["January", "Case study / special stories", "Source report"],
            ["February", "Grammar", "Source report"],
            ["March", "Parents’ involvement in children’s education", "Source report"],
            ["May", "Science experiments, Balwadi", "Source report"],
            ["June", "Records and survey", "Source report"],
            ["July", "Language and numeracy tools; work-life balance", "Source report"],
            ["August", "Class control and library class planning; emotional regulation", "Source report"],
            ["September", "Working with hyperactive/restless children; social skills", "Source report"],
            ["October", "Volunteer management", "Source report"],
            ["November", "Teacher role and responsibilities; parent counselling", "Source report"],
            ["December", "Time management", "Source report"],
        ]
    add_table(doc, ["Month", "Training Topic", "No. of participants"], month_rows, font=10.0)
    page_break(doc)

    # PAGE 13 — photograph page 1
    add_title(doc, "Programme Photographs")
    photo_grid(doc, [
        "Real activity photo – 1",
        "Real activity photo – 2",
        "Real activity photo – 3"
    ], height=2.15)
    page_break(doc)

    # PAGE 14 — photograph page 2
    add_title(doc, "Programme Photographs")
    photo_grid(doc, [
        "Real activity photo – 4",
        "Real activity photo – 5",
        "Real activity photo – 6"
    ], height=2.15)

    out = OUTPUT_DIR / f"{funder}_Progress_Report_V11.docx"
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    return out


def main():
    global INPUT_XLSX, TRANSLATED_JSON
    # Re-discover on every run so the generator picks the cleaned dataset when available.
    candidates=[]
    for pat in ("*CLEANED*.xlsx","MONTH*2025*.xlsx","*.xlsx"):
        candidates.extend(ROOT.rglob(pat))
    candidates=[x for x in candidates if "AUDIT" not in x.name.upper() and not x.name.startswith("~$")]
    candidates=sorted(set(candidates), key=lambda x:("CLEANED" not in x.name.upper(), len(str(x))))
    if candidates:
        INPUT_XLSX=candidates[0]
    elif INPUT_XLSX is None or not INPUT_XLSX.exists():
        raise FileNotFoundError(f"No Excel workbook found under project root: {ROOT}")
    if TRANSLATED_JSON is None or not TRANSLATED_JSON.exists():
        candidates=list(ROOT.rglob("translated_report.json"))
        if candidates: TRANSLATED_JSON=candidates[0]
    all_df=load_data()
    translated=load_translation()
    OUTPUT_DIR.mkdir(parents=True,exist_ok=True)
    for funder in sorted(all_df["FUNDING WISE"].dropna().unique()):
        try:
            out=build_report(str(funder),all_df,translated)
            print(f"Created: {out}")
        except Exception as e:
            print(f"ERROR for {funder}: {e}")


if __name__ == "__main__":
    main()
