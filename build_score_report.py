#!/usr/bin/env python3
"""Create a journalist-friendly Markdown or Word report from Michigan assessment extracts.

Usage:
    python build_score_report.py
    python build_score_report.py --input-dir outputs --output test_score_report.md
    python build_score_report.py --format docx --output test_score_report.docx
    python build_score_report.py --school-year 2025 --min-assessed 10

The script is based on analyze_scores.py but replaces notebook display/print output
with one organized Markdown or Word report. It tolerates common suppression markers and
supports both wide SAT files and long subject-level extracts.
"""

from __future__ import annotations

import argparse
from pathlib import Path
import re
import sys
from typing import Iterable

import numpy as np
import pandas as pd

try:
    from docx import Document
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches, Pt, RGBColor
except ImportError:  # Only required when --format docx is used.
    Document = None

SUPPRESSED = {"", "*", "**", "--", "---", "—", "N/A", "NA", "<5", "< 5", "nan", "None"}
RACE_GROUPS = [
    "American Indian or Alaska Native", "Asian", "Black or African American",
    "Hispanic of Any Race", "Native Hawaiian or Other Pacific Islander",
    "Two or More Races", "White",
]
CATEGORY_GROUPS = {
    "Race and ethnicity": RACE_GROUPS,
    "Economic status": ["Economically Disadvantaged", "Not Economically Disadvantaged"],
    "Gender": ["Female", "Male"],
    "English learner status": ["English Learners", "Not English Learners"],
    "Disability status": ["Students With Disabilities", "Students Without Disabilities"],
}


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--input-dir", type=Path, default=Path("outputs"), help="Directory containing extract CSVs")
    p.add_argument("--output", type=Path, default=Path("test_score_report.md"),
                   help="Output path; suffix is adjusted to match --format")
    p.add_argument("--format", choices=["md", "docx"], default="md",
                   help="Output format: Markdown or Word (default: md)")
    p.add_argument("--school-year", default=None, help="Optional SchoolYear value to keep, such as 2025")
    p.add_argument("--min-assessed", type=int, default=None,
                   help="Optional minimum NumberAssessed for district rankings; recommended when available")
    p.add_argument("--top-n", type=int, default=10, help="Number of districts in rankings (default: 10)")
    return p.parse_args()


def read_csv(input_dir: Path, name: str, required: bool = True) -> pd.DataFrame:
    path = input_dir / name
    if not path.exists():
        if required:
            raise FileNotFoundError(f"Required input not found: {path}")
        return pd.DataFrame()
    df = pd.read_csv(path, dtype=str, low_memory=False)
    df.columns = df.columns.str.strip()
    return df


def norm(value) -> str:
    return "" if pd.isna(value) else str(value).strip().casefold()


def numeric(series: pd.Series) -> pd.Series:
    cleaned = (series.astype("string").str.strip().str.replace(",", "", regex=False)
               .str.replace("%", "", regex=False).replace(list(SUPPRESSED), pd.NA))
    return pd.to_numeric(cleaned, errors="coerce")


def choose_year(df: pd.DataFrame, requested: str | None) -> tuple[pd.DataFrame, str | None]:
    if df.empty or "SchoolYear" not in df:
        return df.copy(), requested
    years = df["SchoolYear"].dropna().astype(str).str.strip()
    if requested is not None:
        return df.loc[years.reindex(df.index).eq(str(requested))].copy(), str(requested)
    if years.empty:
        return df.copy(), None
    # Prefer numeric maximum, otherwise lexical maximum.
    nums = pd.to_numeric(years, errors="coerce")
    selected = years.loc[nums.idxmax()] if nums.notna().any() else sorted(years.unique())[-1]
    return df.loc[df["SchoolYear"].astype(str).str.strip().eq(str(selected))].copy(), str(selected)


def grade_mask(df: pd.DataFrame, grade: int) -> pd.Series:
    if "GradeContentTested" not in df:
        return pd.Series(False, index=df.index)
    s = df["GradeContentTested"].astype("string").fillna("").str.strip()
    return s.str.contains(rf"(?<!\d)0?{grade}(?!\d)", regex=True, na=False)


def subject_mask(df: pd.DataFrame, subject: str) -> pd.Series:
    if "Subject" not in df:
        return pd.Series(False, index=df.index)
    patterns = {
        "ELA": r"^(ELA|English Language Arts|Reading)$",
        "Mathematics": r"^(Mathematics|Math)$",
        "Science": r"^Science$",
        "Social Studies": r"^(Social Studies|Social Science)$",
    }
    return df["Subject"].astype("string").str.strip().str.match(patterns[subject], case=False, na=False)


def mstep_psat(df: pd.DataFrame) -> pd.DataFrame:
    if "TestType" not in df:
        return df.copy()
    return df.loc[df["TestType"].astype("string").str.strip().str.casefold().isin({"m-step", "psat"})].copy()


def metric_series(df: pd.DataFrame, preferred: Iterable[str]) -> tuple[pd.Series, str | None]:
    """Return first available numeric metric. Sum advanced + proficient when needed."""
    for col in preferred:
        if col in df:
            return numeric(df[col]), col
    if {"PercentAdvanced", "PercentProficient"}.issubset(df.columns):
        a, p = numeric(df["PercentAdvanced"]), numeric(df["PercentProficient"])
        result = a.fillna(0) + p.fillna(0)
        result[a.isna() & p.isna()] = np.nan
        return result, "PercentAdvanced + PercentProficient"
    return pd.Series(np.nan, index=df.index, dtype=float), None


def fmt_num(value, decimals: int = 1) -> str:
    return "Not reported" if pd.isna(value) else f"{float(value):,.{decimals}f}"


def fmt_pct(value) -> str:
    return "Not reported" if pd.isna(value) else f"{float(value):.1f}%"


def md_escape(value) -> str:
    return str(value).replace("|", "\\|").replace("\n", " ").strip()


def markdown_table(rows: list[list], headers: list[str], align_right: set[int] | None = None) -> str:
    if not rows:
        return "_No matching records were found._\n"
    align_right = align_right or set()
    lines = ["| " + " | ".join(headers) + " |"]
    lines.append("| " + " | ".join("---:" if i in align_right else "---" for i in range(len(headers))) + " |")
    for row in rows:
        lines.append("| " + " | ".join(md_escape(x) for x in row) + " |")
    return "\n".join(lines) + "\n"


def single_value(df: pd.DataFrame, grade: int, subject: str, metric: pd.Series) -> float:
    mask = grade_mask(df, grade) & subject_mask(df, subject)
    vals = metric.loc[mask].dropna()
    return vals.iloc[0] if not vals.empty else np.nan


def proficiency_table(df: pd.DataFrame) -> tuple[str, str | None]:
    df = mstep_psat(df)
    metric, metric_name = metric_series(df, ["PercentMet", "proficient_or_above_pct"])
    rows = []
    for grade in range(3, 9):
        for subject in ["ELA", "Mathematics", "Science", "Social Studies"]:
            value = single_value(df, grade, subject, metric)
            if not pd.isna(value):
                rows.append([grade, subject, fmt_pct(value)])
    return markdown_table(rows, ["Grade", "Subject", "Met proficiency"], {0, 2}), metric_name


def equity_sections(df: pd.DataFrame) -> list[str]:
    df = mstep_psat(df)
    if "ReportCategory" not in df:
        return ["_The file does not contain a `ReportCategory` column._\n"]
    metric, _ = metric_series(df, ["PercentMet", "proficient_or_above_pct"])
    sections = []
    for heading, groups in CATEGORY_GROUPS.items():
        rows = []
        for group in groups:
            gmask = df["ReportCategory"].astype("string").str.strip().eq(group)
            ela3 = metric.loc[gmask & grade_mask(df, 3) & subject_mask(df, "ELA")].dropna()
            math8 = metric.loc[gmask & grade_mask(df, 8) & subject_mask(df, "Mathematics")].dropna()
            if not ela3.empty or not math8.empty:
                rows.append([group, fmt_pct(ela3.iloc[0] if not ela3.empty else np.nan),
                             fmt_pct(math8.iloc[0] if not math8.empty else np.nan)])
        sections.extend([f"### {heading}\n", markdown_table(rows,
                         ["Student group", "Grade 3 ELA", "Grade 8 math"], {1, 2})])
    return sections


def generic_psat_table(df: pd.DataFrame) -> str:
    if df.empty:
        return "_No PSAT extract was available._\n"
    met, _ = metric_series(df, ["PercentMet", "proficient_or_above_pct"])
    avg, _ = metric_series(df, ["AvgSS", "AllSubjectScoreAverage", "MathScoreAverage", "EBRWScoreAverage"])
    rows = []
    for i, row in df.iterrows():
        subject = row.get("Subject", "All subjects")
        rows.append([subject, fmt_num(row.get("NumberAssessed", np.nan), 0), fmt_num(avg.loc[i]), fmt_pct(met.loc[i])])
    return markdown_table(rows, ["Subject", "Students assessed", "Average score", "Met benchmark"], {1, 2, 3})


def sat_summary(df: pd.DataFrame) -> tuple[str, list[str]]:
    if df.empty:
        return "_No SAT records were found._\n", []
    if "Subgroup" in df:
        all_students = df.loc[df["Subgroup"].astype("string").str.strip().str.casefold().eq("all students")]
        if not all_students.empty:
            df = all_students
    notes = []
    rows = []
    # Wide schema used by analyze_scores.py.
    if any(c in df for c in ["AllSubjectScoreAverage", "MathScoreAverage", "EBRWScoreAverage"]):
        r = df.iloc[0]
        assessed = r.get("MathNumAssessed", r.get("NumberAssessed", np.nan))
        for label, score_col, ready_col in [
            ("Overall", "AllSubjectScoreAverage", "AllSubjectPercentReady"),
            ("Math", "MathScoreAverage", "MathPercentReady"),
            ("Evidence-based reading and writing", "EBRWScoreAverage", "EBRWPercentReady"),
        ]:
            if score_col in df or ready_col in df:
                rows.append([label, fmt_num(assessed, 0), fmt_num(pd.to_numeric(pd.Series([r.get(score_col)]), errors="coerce").iloc[0]),
                             fmt_pct(pd.to_numeric(pd.Series([r.get(ready_col)]), errors="coerce").iloc[0])])
    else:
        avg, _ = metric_series(df, ["AvgSS", "MeanPtsEarned"])
        ready, _ = metric_series(df, ["PercentMet", "proficient_or_above_pct"])
        for i, r in df.iterrows():
            rows.append([r.get("Subject", "Overall"), fmt_num(pd.to_numeric(pd.Series([r.get("NumberAssessed")]), errors="coerce").iloc[0], 0),
                         fmt_num(avg.loc[i]), fmt_pct(ready.loc[i])])
    return markdown_table(rows, ["SAT measure", "Students assessed", "Average score", "College ready / met benchmark"], {1, 2, 3}), notes


def district_rankings(df: pd.DataFrame, score_col: str, label: str, top_n: int,
                      min_assessed: int | None = None, filters: pd.Series | None = None,
                      percent: bool = False) -> str:
    if df.empty or score_col not in df or "DistrictName" not in df:
        return f"_Unable to rank districts for {label}: required fields were not available._\n"
    work = df.copy()
    if filters is not None:
        work = work.loc[filters.reindex(work.index, fill_value=False)].copy()
    work["_value"] = numeric(work[score_col])
    if min_assessed is not None and "NumberAssessed" in work:
        work = work.loc[numeric(work["NumberAssessed"]) >= min_assessed]
    work = work.dropna(subset=["_value", "DistrictName"]).sort_values("_value", ascending=False).head(top_n)
    rows = [[rank, r["DistrictName"], fmt_pct(r["_value"]) if percent else fmt_num(r["_value"])]
            for rank, (_, r) in enumerate(work.iterrows(), 1)]
    return markdown_table(rows, ["Rank", "District", label], {0, 2})



def _set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def _set_cell_margins(cell, top=90, start=110, bottom=90, end=110) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for name, value in {"top": top, "start": start, "bottom": bottom, "end": end}.items():
        node = tc_mar.find(qn(f"w:{name}"))
        if node is None:
            node = OxmlElement(f"w:{name}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def _repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def _add_page_number(paragraph) -> None:
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instruction, separate, text, end])


def _configure_document(document) -> None:
    section = document.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    normal = document.styles["Normal"]
    normal.font.name = "Aptos"
    normal.font.size = Pt(10)
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.line_spacing = 1.05

    title = document.styles["Title"]
    title.font.name = "Aptos Display"
    title.font.size = Pt(24)
    title.font.bold = True
    title.font.color.rgb = RGBColor(31, 78, 121)

    for style_name, color in {
        "Heading 1": RGBColor(31, 78, 121),
        "Heading 2": RGBColor(47, 84, 150),
        "Heading 3": RGBColor(68, 68, 68),
    }.items():
        style = document.styles[style_name]
        style.font.name = "Aptos Display"
        style.font.color.rgb = color
        style.font.bold = True
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.space_before = Pt(12)
        style.paragraph_format.space_after = Pt(5)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run("Michigan state test score report  |  Page ")
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(100, 100, 100)
    _add_page_number(footer)


def _add_inline_markdown(paragraph, text: str) -> None:
    """Render the report's limited inline Markdown as native Word runs."""
    # Internal Markdown navigation links are unnecessary in Word; retain the label.
    text = re.sub(r"\[([^]]+)\]\(#[^)]+\)", r"\1", text)
    token_re = re.compile(r"(\*\*.*?\*\*|`.*?`|“.*?”)")
    position = 0
    for match in token_re.finditer(text):
        if match.start() > position:
            paragraph.add_run(text[position:match.start()])
        token = match.group(0)
        if token.startswith("**"):
            run = paragraph.add_run(token[2:-2])
            run.bold = True
        elif token.startswith("`"):
            run = paragraph.add_run(token[1:-1])
            run.font.name = "Consolas"
            run.font.size = Pt(9)
        else:
            paragraph.add_run(token)
        position = match.end()
    if position < len(text):
        paragraph.add_run(text[position:])


def _parse_markdown_table(lines: list[str]) -> tuple[list[str], list[list[str]], set[int]]:
    def split_row(line: str) -> list[str]:
        return [cell.strip().replace("\\|", "|") for cell in line.strip().strip("|").split("|")]
    headers = split_row(lines[0])
    alignment = split_row(lines[1]) if len(lines) > 1 else []
    numeric_columns = {i for i, cell in enumerate(alignment) if cell.endswith(":")}
    rows = [split_row(line) for line in lines[2:]]
    return headers, rows, numeric_columns


def _add_word_table(document, headers: list[str], rows: list[list[str]], numeric_columns: set[int]) -> None:
    table = document.add_table(rows=1, cols=len(headers))
    table.style = "Light Shading Accent 1"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    header_row = table.rows[0]
    _repeat_table_header(header_row)

    for index, header in enumerate(headers):
        cell = header_row.cells[index]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        _set_cell_shading(cell, "1F4E79")
        _set_cell_margins(cell)
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT if index in numeric_columns else WD_ALIGN_PARAGRAPH.LEFT
        run = paragraph.add_run(header)
        run.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.size = Pt(9)

    for row_number, values in enumerate(rows):
        cells = table.add_row().cells
        for index, cell in enumerate(cells):
            value = values[index] if index < len(values) else ""
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            _set_cell_margins(cell)
            if row_number % 2 == 1:
                _set_cell_shading(cell, "F2F6FA")
            paragraph = cell.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT if index in numeric_columns else WD_ALIGN_PARAGRAPH.LEFT
            run = paragraph.add_run(value)
            run.font.size = Pt(9)
    document.add_paragraph()


def markdown_to_docx(markdown_text: str, output_path: Path) -> None:
    """Convert this script's structured Markdown into a polished Word report."""
    if Document is None:
        raise RuntimeError(
            "Word output requires python-docx. Install it with: pip install python-docx"
        )

    document = Document()
    _configure_document(document)
    lines = markdown_text.splitlines()
    index = 0
    first_h1 = True

    while index < len(lines):
        raw = lines[index]
        line = raw.strip()
        if not line:
            index += 1
            continue

        if line.startswith("| "):
            table_lines = []
            while index < len(lines) and lines[index].strip().startswith("|"):
                table_lines.append(lines[index].strip())
                index += 1
            if len(table_lines) >= 2:
                headers, rows, numeric_columns = _parse_markdown_table(table_lines)
                _add_word_table(document, headers, rows, numeric_columns)
            continue

        if line.startswith("# "):
            text = line[2:].strip()
            if first_h1:
                paragraph = document.add_paragraph(style="Title")
                first_h1 = False
            else:
                paragraph = document.add_heading(level=1)
            _add_inline_markdown(paragraph, text)
        elif line.startswith("## "):
            paragraph = document.add_heading(level=1)
            _add_inline_markdown(paragraph, line[3:].strip())
        elif line.startswith("### "):
            paragraph = document.add_heading(level=2)
            _add_inline_markdown(paragraph, line[4:].strip())
        elif line.startswith("> "):
            table = document.add_table(rows=1, cols=1)
            table.style = "Light Shading Accent 1"
            cell = table.cell(0, 0)
            _set_cell_shading(cell, "EAF2F8")
            _set_cell_margins(cell, top=150, start=170, bottom=150, end=170)
            paragraph = cell.paragraphs[0]
            _add_inline_markdown(paragraph, line[2:].strip())
            document.add_paragraph()
        elif line.startswith("- "):
            paragraph = document.add_paragraph(style="List Bullet")
            _add_inline_markdown(paragraph, line[2:].strip())
        elif line.startswith("_The file") or (line.startswith("_") and line.endswith("_")):
            paragraph = document.add_paragraph()
            run = paragraph.add_run(line.strip("_"))
            run.italic = True
            run.font.color.rgb = RGBColor(95, 95, 95)
        else:
            paragraph = document.add_paragraph()
            _add_inline_markdown(paragraph, line.replace("  ", " "))
        index += 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)

def main() -> int:
    args = parse_args()
    try:
        files = {
            "state_g38": read_csv(args.input_dir, "grades_3_8_proficiency_statewide.csv"),
            "state_equity": read_csv(args.input_dir, "grade_subject_report_categories_statewide.csv"),
            "state_psat": read_csv(args.input_dir, "psat8_statewide.csv", required=False),
            "state_sat": read_csv(args.input_dir, "sat_statewide.csv"),
            "detroit_g38": read_csv(args.input_dir, "grades_3_8_proficiency_dpscd.csv"),
            "detroit_equity": read_csv(args.input_dir, "grade_subject_report_categories_dpscd.csv", required=False),
            "detroit_psat": read_csv(args.input_dir, "psat8_dpscd.csv"),
            "detroit_sat": read_csv(args.input_dir, "sat_dpscd.csv"),
            "district_g38": read_csv(args.input_dir, "grades_3_8_proficiency_all_districts.csv"),
            "district_sat": read_csv(args.input_dir, "sat_all_districts.csv"),
        }
    except (FileNotFoundError, pd.errors.ParserError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    selected_years = []
    for key, df in files.items():
        files[key], year = choose_year(df, args.school_year)
        if year:
            selected_years.append(year)
    years = sorted(set(selected_years))
    year_label = ", ".join(years) if years else "year not specified in source files"

    lines = [
        "# Michigan state test score results\n",
        f"**School year:** {year_label}  \n",
        "**Geographies:** Michigan statewide, Detroit Public Schools Community District and district rankings  \n",
        "\n> **How to read this report:** Percentages marked as proficiency or benchmark attainment are reproduced from the assessment extracts. "
        "Suppressed or unavailable values appear as “Not reported.” District rankings describe reported results only and should not be treated "
        "as measures of district quality. Enrollment size, student population and suppression rules can affect comparisons.\n",
        "## Quick navigation\n",
        "- [Statewide grades 3-8 proficiency](#statewide-grades-3-8-proficiency)\n",
        "- [Statewide results by student group](#statewide-results-by-student-group)\n",
        "- [Statewide eighth-grade PSAT](#statewide-eighth-grade-psat)\n",
        "- [Statewide SAT](#statewide-sat)\n",
        "- [Detroit results](#detroit-public-schools-community-district)\n",
        "- [Top district results](#top-district-results)\n",
    ]

    table, metric_name = proficiency_table(files["state_g38"])
    lines += ["\n## Statewide grades 3-8 proficiency\n", table]
    if metric_name:
        lines += [f"_Metric used: `{metric_name}`._\n"]

    lines += ["\n## Statewide results by student group\n",
              "These tables isolate third-grade ELA and eighth-grade mathematics, matching the equity outputs in the analysis notebook.\n"]
    lines += equity_sections(files["state_equity"])

    lines += ["\n## Statewide eighth-grade PSAT\n", generic_psat_table(files["state_psat"])]
    sat_table, _ = sat_summary(files["state_sat"])
    lines += ["\n## Statewide SAT\n", sat_table]

    lines += ["\n## Detroit Public Schools Community District\n", "### Grades 3-8 proficiency\n"]
    detroit_table, _ = proficiency_table(files["detroit_g38"])
    lines += [detroit_table, "### Third-grade ELA and eighth-grade math by student group\n"]
    lines += equity_sections(files["detroit_equity"])
    lines += ["### Eighth-grade PSAT\n", generic_psat_table(files["detroit_psat"])]
    det_sat_table, _ = sat_summary(files["detroit_sat"])
    lines += ["### SAT\n", det_sat_table]

    lines += ["\n## Top district results\n"]
    if args.min_assessed is None:
        lines += ["> **Ranking caution:** No minimum tested-student threshold was applied. Small districts may appear near the top, and suppressed values are excluded.\n"]
    else:
        lines += [f"> Rankings include only records with at least **{args.min_assessed:,} students assessed**, when that field is available.\n"]

    sat_df = files["district_sat"]
    if "AllSubjectScoreAverage" in sat_df:
        lines += ["### Highest average SAT total scores\n",
                  district_rankings(sat_df, "AllSubjectScoreAverage", "Average SAT total score", args.top_n, args.min_assessed)]
        lines += ["### Highest average SAT math scores\n",
                  district_rankings(sat_df, "MathScoreAverage", "Average SAT math score", args.top_n, args.min_assessed)]
        lines += ["### Highest average SAT reading and writing scores\n",
                  district_rankings(sat_df, "EBRWScoreAverage", "Average SAT EBRW score", args.top_n, args.min_assessed)]
    else:
        score, score_name = metric_series(sat_df, ["AvgSS", "MeanPtsEarned"])
        sat_df = sat_df.copy(); sat_df["_score"] = score
        subject = sat_df.get("Subject", pd.Series("", index=sat_df.index)).astype("string")
        lines += ["### Highest average SAT overall scores\n",
                  district_rankings(sat_df, "_score", "Average SAT score", args.top_n, args.min_assessed,
                                    subject.str.contains(r"overall|total|all", case=False, regex=True, na=False))]
        lines += ["### Highest average SAT math scores\n",
                  district_rankings(sat_df, "_score", "Average SAT math score", args.top_n, args.min_assessed,
                                    subject.str.contains("math", case=False, na=False))]
        lines += ["### Highest average SAT reading and writing scores\n",
                  district_rankings(sat_df, "_score", "Average SAT EBRW score", args.top_n, args.min_assessed,
                                    subject.str.contains(r"reading|writing|ebrw|ela", case=False, regex=True, na=False))]

    prof = mstep_psat(files["district_g38"])
    prof_metric, prof_col = metric_series(prof, ["PercentMet", "proficient_or_above_pct"])
    prof = prof.copy(); prof["_proficiency"] = prof_metric
    lines += ["### Highest third-grade ELA proficiency\n",
              district_rankings(prof, "_proficiency", "Met proficiency", args.top_n, args.min_assessed,
                                grade_mask(prof, 3) & subject_mask(prof, "ELA"), percent=True)]
    lines += ["### Highest eighth-grade math proficiency\n",
              district_rankings(prof, "_proficiency", "Met proficiency", args.top_n, args.min_assessed,
                                grade_mask(prof, 8) & subject_mask(prof, "Mathematics"), percent=True)]

    lines += ["\n## Reporting notes\n",
              "- “Not reported” means the value was absent, suppressed or could not be parsed as a number.\n",
              "- Rankings exclude suppressed and nonnumeric values. Ties are left in source-file order.\n",
              "- Confirm unusually high or low district values against the original state release before publication.\n",
              "- If the input directory contains multiple school years, the script defaults to the latest year found. Use `--school-year` to select one explicitly.\n"]

    report_text = "\n".join(lines)
    output_path = args.output.with_suffix(".docx" if args.format == "docx" else ".md")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if args.format == "docx":
        try:
            markdown_to_docx(report_text, output_path)
        except RuntimeError as exc:
            print(f"ERROR: {exc}", file=sys.stderr)
            return 1
        print(f"Wrote Word report: {output_path.resolve()}")
    else:
        output_path.write_text(report_text, encoding="utf-8")
        print(f"Wrote Markdown report: {output_path.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
