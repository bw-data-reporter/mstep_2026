#!/usr/bin/env python3
"""Create a journalist-friendly Markdown report from Michigan assessment extracts.

Usage:
    python build_score_report.py
    python build_score_report.py --input-dir outputs --output test_score_report.md
    python build_score_report.py --school-year 2025 --min-assessed 10

The script is based on analyze_scores.py but replaces notebook display/print output
with one organized Markdown report. It tolerates common suppression markers and
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


#for PDF handling
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

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
    p.add_argument("--output", type=Path, default=Path("test_score_report.md"), help="Markdown output path")
    p.add_argument("--school-year", default=None, help="Optional SchoolYear value to keep, such as 2025")
    p.add_argument("--min-assessed", type=int, default=None,
                   help="Optional minimum NumberAssessed for district rankings; recommended when available")
    p.add_argument("--top-n", type=int, default=10, help="Number of districts in rankings (default: 10)")
    p.add_argument("--format", choices=["md", "pdf"], default="md", help="Output format (default: md)")
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

def markdown_to_pdf(markdown_text: str, pdf_path: Path) -> None:
    """
    Convert the generated markdown report into a simple,
    journalist-friendly PDF.
    """
    doc = SimpleDocTemplate(str(pdf_path))

    styles = getSampleStyleSheet()
    body = styles["BodyText"]
    heading1 = styles["Heading1"]
    heading2 = styles["Heading2"]
    heading3 = styles["Heading3"]

    elements = []

    for line in markdown_text.splitlines():
        line = line.strip()

        if not line:
            elements.append(Spacer(1, 6))
            continue

        if line.startswith("# "):
            elements.append(Paragraph(line[2:], heading1))
            continue

        if line.startswith("## "):
            elements.append(Paragraph(line[3:], heading2))
            continue

        if line.startswith("### "):
            elements.append(Paragraph(line[4:], heading3))
            continue

        # Render markdown tables as monospaced text blocks.
        if line.startswith("|"):
            elements.append(
                Paragraph(
                    f"<font face='Courier'>{line}</font>",
                    body
                )
            )
            continue

        elements.append(Paragraph(line, body))

    doc.build(elements)


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
        "\n> **How to read this report:** Percentages marked as proficiency or benchmark attainment are reproduced from the assessment extracts.",
        "Suppressed or unavailable values appear as “Not reported. \n",
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
    ]


    # args.output.parent.mkdir(parents=True, exist_ok=True)
    # args.output.write_text("\n".join(lines), encoding="utf-8")
    # print(f"Wrote Markdown report: {args.output.resolve()}")
    report_text = "\n".join(lines)

    args.output.parent.mkdir(parents=True, exist_ok=True)

    if args.format == "md":
        output_path = args.output

        if output_path.suffix.lower() != ".md":
            output_path = output_path.with_suffix(".md")

        output_path.write_text(report_text, encoding="utf-8")
        print(f"Wrote Markdown report: {output_path.resolve()}")

    else:
        output_path = args.output

        if output_path.suffix.lower() != ".pdf":
            output_path = output_path.with_suffix(".pdf")

        markdown_to_pdf(report_text, output_path)
        print(f"Wrote PDF report: {output_path.resolve()}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
