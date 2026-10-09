from datetime import date, datetime
from html import escape
from io import BytesIO
import os
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qs, urlencode, urlsplit
from urllib.request import HTTPRedirectHandler, Request, build_opener
from zipfile import BadZipFile

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import reportlab
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

st.set_page_config(page_title="Recruitment Profiles Dashboard", layout="wide")

SHEET_NAME = "Master Sheet"
REQUIRED_COLUMNS = ("Date", "Consultant", "Job Title", "Status")
REFRESH_SECONDS = 10
DEFAULT_CHART_COLORS = (
    "#0068c9",
    "#83c9ff",
    "#ff2b2b",
    "#ffabab",
    "#29b09d",
    "#7defa1",
    "#6d3fc0",
    "#d5dae5",
    "#ff8700",
    "#ffd16a",
    "#1c83e1",
    "#a3c4dc",
)


def _google_drive_file_id(url: str) -> tuple[str, bool]:
    parsed_url = urlsplit(url)
    hostname = (parsed_url.hostname or "").lower()
    if parsed_url.scheme != "https" or hostname not in {
        "drive.google.com",
        "docs.google.com",
    }:
        raise ValueError(
            "Configure an HTTPS sharing link from Google Drive or Google Sheets."
        )

    path_parts = parsed_url.path.strip("/").split("/")
    is_google_sheet = hostname == "docs.google.com" and "spreadsheets" in path_parts
    file_id = ""
    if "d" in path_parts:
        file_id = path_parts[path_parts.index("d") + 1]
    else:
        file_id = parse_qs(parsed_url.query).get("id", [""])[0]
    if not file_id or not all(
        character.isalnum() or character in "_-" for character in file_id
    ):
        raise ValueError(
            "Could not find a Google Drive file ID in that link. Copy the "
            "workbook's sharing link from Google Drive."
        )
    return file_id, is_google_sheet


class _GoogleDriveRedirectHandler(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        parsed_url = urlsplit(newurl)
        hostname = (parsed_url.hostname or "").lower()
        allowed_hosts = (
            hostname in {"drive.google.com", "docs.google.com"}
            or hostname.endswith(".googleusercontent.com")
        )
        if parsed_url.scheme != "https" or not allowed_hosts:
            raise ValueError(
                "Google Drive redirected the workbook download to an "
                "unsupported host."
            )
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def download_google_drive_workbook(url: str) -> bytes:
    """Fetch an Excel workbook shared from Google Drive or Google Sheets."""
    file_id, is_google_sheet = _google_drive_file_id(url)
    if is_google_sheet:
        download_url = (
            f"https://docs.google.com/spreadsheets/d/{file_id}/export?"
            f"{urlencode({'format': 'xlsx'})}"
        )
    else:
        download_url = (
            "https://drive.google.com/uc?"
            f"{urlencode({'export': 'download', 'id': file_id})}"
        )
    request = Request(
        download_url,
        headers={"User-Agent": "Mozilla/5.0", "Accept": "*/*"},
    )
    try:
        with build_opener(_GoogleDriveRedirectHandler).open(
            request, timeout=30
        ) as response:
            file_contents = response.read()
    except HTTPError as error:
        raise RuntimeError(
            f"Google Drive returned HTTP {error.code}. Check that the workbook "
            "is shared with link access."
        ) from error
    if not file_contents.startswith(b"PK"):
        raise ValueError(
            "Google Drive did not return an Excel workbook. Set the link to "
            "Anyone with the link — Viewer, and make sure it opens without "
            "asking the dashboard to sign in."
        )
    return file_contents


@st.cache_data(show_spinner="Loading the Master Sheet...")
def load_profiles(file_contents: bytes) -> pd.DataFrame:
    """Load and normalize the dashboard fields from the Master Sheet."""
    profiles = pd.read_excel(
        BytesIO(file_contents), sheet_name=SHEET_NAME, engine="openpyxl"
    )

    missing_columns = [column for column in REQUIRED_COLUMNS if column not in profiles]
    if missing_columns:
        raise ValueError(
            f"The {SHEET_NAME!r} sheet is missing required columns: "
            f"{', '.join(missing_columns)}"
        )

    profiles = profiles.dropna(subset=list(REQUIRED_COLUMNS), how="all").copy()
    for column in ("Consultant", "Job Title", "Status"):
        profiles[column] = (
            profiles[column]
            .astype("string")
            .str.strip()
            .replace("", pd.NA)
            .fillna("Not specified")
        )
    profiles["Date"] = pd.to_datetime(
        profiles["Date"], errors="coerce", dayfirst=True
    )
    return profiles


def profile_counts(profiles: pd.DataFrame, column: str) -> pd.DataFrame:
    return (
        profiles.groupby(column, dropna=False)
        .size()
        .reset_index(name="Profiles")
        .sort_values("Profiles", ascending=False)
    )


def create_dashboard_charts(profiles: pd.DataFrame) -> dict[str, go.Figure]:
    chart_colors = (
        st.get_option("theme.chartCategoricalColors") or DEFAULT_CHART_COLORS
    )
    consultant_counts = profile_counts(profiles, "Consultant")
    title_counts = profile_counts(profiles, "Job Title")
    consultant_title_counts = (
        profiles.groupby(["Consultant", "Job Title"], dropna=False)
        .size()
        .reset_index(name="Profiles")
    )
    status_counts = profile_counts(profiles, "Status")

    dated_profiles = profiles.dropna(subset=["Date"])
    daily_counts = (
        dated_profiles.groupby(["Date", "Consultant"])
        .size()
        .reset_index(name="Profiles")
        .sort_values("Date")
    )

    status_chart = px.pie(
        status_counts,
        names="Status",
        values="Profiles",
        hole=0.45,
        color_discrete_sequence=chart_colors,
    )
    status_chart.update_traces(textinfo="label+percent")

    return {
        "Profiles by Consultant": px.bar(
            consultant_counts,
            x="Consultant",
            y="Profiles",
            color="Consultant",
            color_discrete_sequence=chart_colors,
        ),
        "Profiles by Job Title": px.bar(
            title_counts,
            x="Job Title",
            y="Profiles",
            color="Job Title",
            color_discrete_sequence=chart_colors,
        ),
        "Profiles by Consultant and Job Title": px.bar(
            consultant_title_counts,
            x="Consultant",
            y="Profiles",
            color="Job Title",
            barmode="group",
            labels={"Profiles": "Count"},
            color_discrete_sequence=chart_colors,
        ),
        "Profiles by Status": status_chart,
        "Profiles Shared by Date and Consultant": px.line(
            daily_counts,
            x="Date",
            y="Profiles",
            color="Consultant",
            markers=True,
            color_discrete_sequence=chart_colors,
        ),
    }


@st.cache_data(show_spinner="Preparing the PDF report...")
def build_pdf_report(profiles: pd.DataFrame, filter_summary: str) -> bytes:
    """Create a paginated PDF report for the currently filtered profiles."""
    buffer = BytesIO()
    page_size = letter
    document = SimpleDocTemplate(
        buffer,
        pagesize=page_size,
        rightMargin=0.5 * inch,
        leftMargin=0.5 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch,
        title="Recruitment Profiles Report",
    )
    font_dir = Path(reportlab.__file__).parent / "fonts"
    pdfmetrics.registerFont(TTFont("DashboardVera", str(font_dir / "Vera.ttf")))
    pdfmetrics.registerFont(TTFont("DashboardVera-Bold", str(font_dir / "VeraBd.ttf")))
    pdfmetrics.registerFontFamily(
        "DashboardVera",
        normal="DashboardVera",
        bold="DashboardVera-Bold",
        italic="DashboardVera",
        boldItalic="DashboardVera-Bold",
    )
    styles = getSampleStyleSheet()
    styles["Normal"].fontName = "DashboardVera"
    styles["BodyText"].fontName = "DashboardVera"
    styles["Title"].fontName = "DashboardVera-Bold"
    styles["Heading2"].fontName = "DashboardVera-Bold"
    styles["BodyText"].fontSize = 8
    styles["BodyText"].leading = 10
    styles["Heading2"].spaceBefore = 10
    styles["Heading2"].spaceAfter = 5
    styles.add(
        ParagraphStyle(
            name="ChartTableBody",
            parent=styles["BodyText"],
            fontSize=8,
            leading=9,
        )
    )
    story = [
        Paragraph("Recruitment Profiles Report", styles["Title"]),
        Paragraph(
            f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M')} · "
            f"{escape(filter_summary)}",
            styles["BodyText"],
        ),
        Spacer(1, 10),
    ]

    metrics = [
        ("Profiles shared", f"{len(profiles):,}"),
        ("Consultants", f"{profiles['Consultant'].nunique():,}"),
        ("Job titles", f"{profiles['Job Title'].nunique():,}"),
        ("Statuses", f"{profiles['Status'].nunique():,}"),
    ]
    metric_table = Table(
        [
            [Paragraph(f"<b>{label}</b>", styles["BodyText"]) for label, _ in metrics],
            [Paragraph(value, styles["Heading2"]) for _, value in metrics],
        ],
        colWidths=[(page_size[0] - inch) / 4] * 4,
    )
    metric_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#eef4fb")),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#c8d6e5")),
                ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#c8d6e5")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.extend([metric_table, Spacer(1, 8)])

    chart_width = page_size[0] - inch
    chart_page_content_height = page_size[1] - inch

    def make_chart_data_table(
        first_column: str,
        rows: list[tuple[str, int, float]],
    ) -> Table:
        table_rows = [
            [
                Paragraph(f"<b>{escape(first_column)}</b>", styles["ChartTableBody"]),
                Paragraph("<b>Profiles</b>", styles["ChartTableBody"]),
                Paragraph("<b>Percentage</b>", styles["ChartTableBody"]),
            ]
        ]
        table_rows.extend(
            [
                [
                    Paragraph(escape(label), styles["ChartTableBody"]),
                    Paragraph(f"{count:,}", styles["ChartTableBody"]),
                    Paragraph(f"{percentage:.1%}", styles["ChartTableBody"]),
                ]
                for label, count, percentage in rows
            ]
        )
        if len(table_rows) == 1:
            table_rows.append(
                [
                    Paragraph("No matching data", styles["ChartTableBody"]),
                    Paragraph("0", styles["ChartTableBody"]),
                    Paragraph("0.0%", styles["ChartTableBody"]),
                ]
            )
        table = Table(
            table_rows,
            colWidths=[
                chart_width * 0.68,
                chart_width * 0.16,
                chart_width * 0.16,
            ],
            repeatRows=1,
            hAlign="LEFT",
        )
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dce8f5")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#17365d")),
                    ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#c8d6e5")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                    ("TOPPADDING", (0, 0), (-1, -1), 1),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                ]
            )
        )
        return table

    def add_chart_page(
        title: str,
        figure: go.Figure,
        first_column: str,
        rows: list[tuple[str, int, float]],
        data_table: Table | None = None,
        table_heading: str = "Counts and percentages",
    ) -> None:
        table = (
            data_table
            if data_table is not None
            else make_chart_data_table(first_column, rows)
        )
        _, table_height = table.wrap(chart_width, chart_page_content_height)
        chart_height = min(
            480,
            max(220, chart_page_content_height - 55 - table_height),
        )
        figure_image = figure.to_image(
            format="png",
            width=1200,
            height=round(chart_height * 1200 / chart_width),
        )
        story.extend(
            [
                PageBreak(),
                Paragraph(escape(title), styles["Heading2"]),
                Spacer(1, 8),
                Image(
                    BytesIO(figure_image),
                    width=chart_width,
                    height=chart_height,
                    hAlign="CENTER",
                ),
                Paragraph(f"<b>{escape(table_heading)}</b>", styles["BodyText"]),
                Spacer(1, 4),
                table,
            ]
        )

    chart_figures = create_dashboard_charts(profiles)
    consultant_counts = profile_counts(profiles, "Consultant")
    total_profiles = len(profiles)
    add_chart_page(
        "Profiles by Consultant",
        chart_figures["Profiles by Consultant"],
        "Consultant",
        [
            (
                str(consultant),
                int(count),
                int(count) / total_profiles if total_profiles else 0,
            )
            for consultant, count in consultant_counts.itertuples(
                index=False, name=None
            )
        ],
    )
    title_counts = profile_counts(profiles, "Job Title")
    add_chart_page(
        "Profiles by Job Title",
        chart_figures["Profiles by Job Title"],
        "Job Title",
        [
            (
                str(job_title),
                int(count),
                int(count) / total_profiles if total_profiles else 0,
            )
            for job_title, count in title_counts.itertuples(
                index=False, name=None
            )
        ],
    )

    consultant_title_counts = (
        profiles.groupby(["Consultant", "Job Title"], dropna=False)
        .size()
        .reset_index(name="Profiles")
    )
    consultant_job_title_figure = chart_figures[
        "Profiles by Consultant and Job Title"
    ]
    consultant_job_title_figure.update_traces(
        texttemplate="%{y}",
        textposition="outside",
        cliponaxis=False,
    )
    consultants = sorted(profiles["Consultant"].unique().tolist())
    job_titles = sorted(profiles["Job Title"].unique().tolist())
    consultant_title_matrix = (
        consultant_title_counts.pivot(
            index="Job Title",
            columns="Consultant",
            values="Profiles",
        )
        .reindex(index=job_titles, columns=consultants, fill_value=0)
        .fillna(0)
        .astype(int)
    )
    matrix_rows = [
        [
            Paragraph("<b>Job Title</b>", styles["ChartTableBody"]),
            *[
                Paragraph(f"<b>{escape(str(consultant))}</b>", styles["ChartTableBody"])
                for consultant in consultants
            ],
        ]
    ]
    matrix_rows.extend(
        [
            Paragraph(escape(str(job_title)), styles["ChartTableBody"]),
            *[
                Paragraph(
                    f"{int(consultant_title_matrix.loc[job_title, consultant]):,}",
                    styles["ChartTableBody"],
                )
                for consultant in consultants
            ],
        ]
        for job_title in job_titles
    )
    if not job_titles:
        matrix_rows.append(
            [
                Paragraph("No matching data", styles["ChartTableBody"]),
                *[
                    Paragraph("0", styles["ChartTableBody"])
                    for _ in consultants
                ],
            ]
        )
    matrix_table = Table(
        matrix_rows,
        colWidths=[
            chart_width * 0.3,
            *[
                chart_width * 0.7 / max(len(consultants), 1)
                for _ in consultants
            ],
        ],
        repeatRows=1,
        hAlign="LEFT",
    )
    matrix_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dce8f5")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#17365d")),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#c8d6e5")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 1),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ]
        )
    )
    add_chart_page(
        "Profiles by Consultant and Job Title",
        consultant_job_title_figure,
        "Consultant · Job Title",
        [
            (
                f"{consultant} · {job_title}",
                int(count),
                int(count) / total_profiles if total_profiles else 0,
            )
            for consultant, job_title, count in consultant_title_counts.itertuples(
                index=False, name=None
            )
        ],
        data_table=matrix_table,
        table_heading="Counts by Job Title and Consultant",
    )

    status_counts = profile_counts(profiles, "Status")
    add_chart_page(
        "Profiles by Status",
        chart_figures["Profiles by Status"],
        "Status",
        [
            (
                str(status),
                int(count),
                int(count) / total_profiles if total_profiles else 0,
            )
            for status, count in status_counts.itertuples(
                index=False, name=None
            )
        ],
    )

    dated_profiles = profiles.dropna(subset=["Date"])
    daily_counts = (
        dated_profiles.groupby(["Date", "Consultant"])
        .size()
        .reset_index(name="Profiles")
        .sort_values("Date")
    )
    date_totals = daily_counts.groupby("Date")["Profiles"].sum()
    add_chart_page(
        "Profiles Shared by Date and Consultant",
        chart_figures["Profiles Shared by Date and Consultant"],
        "Date · Consultant",
        [
            (
                f"{pd.Timestamp(timestamp).strftime('%Y-%m-%d')} · {consultant}",
                int(count),
                int(count) / int(date_totals.loc[timestamp])
                if int(date_totals.loc[timestamp])
                else 0,
            )
            for timestamp, consultant, count in daily_counts.itertuples(
                index=False, name=None
            )
        ],
    )

    story.append(PageBreak())

    story.append(Paragraph("Filtered Profile Summary", styles["Heading2"]))
    detail_rows = [
        [
            Paragraph(f"<b>{escape(column)}</b>", styles["BodyText"])
            for column in ("Date", "Consultant", "Job Title Counts", "Status Counts")
        ]
    ]
    sorted_profiles = profiles.sort_values(
        "Date", ascending=False, na_position="last"
    )
    for (timestamp, consultant), group in sorted_profiles.groupby(
        ["Date", "Consultant"], dropna=False, sort=False
    ):
        job_title_counts = group["Job Title"].value_counts().sort_index()
        status_counts = group["Status"].value_counts().sort_index()
        title_summary = ", ".join(
            f"{count:,} {escape(str(job_title))}"
            for job_title, count in job_title_counts.items()
        )
        status_summary = ", ".join(
            f"{count:,} {escape(str(status))}"
            for status, count in status_counts.items()
        )
        detail_rows.append(
            [
                Paragraph(
                    escape(
                        timestamp.strftime("%d.%m.%y") if pd.notna(timestamp) else ""
                    ),
                    styles["BodyText"],
                ),
                Paragraph(escape(str(consultant)), styles["BodyText"]),
                Paragraph(title_summary, styles["BodyText"]),
                Paragraph(status_summary, styles["BodyText"]),
            ]
        )
    if len(detail_rows) == 1:
        detail_rows.append(
            [
                Paragraph("No matching profiles", styles["BodyText"]),
                "",
                "",
                "",
            ]
        )
    detail_table = Table(
        detail_rows,
        colWidths=[70, 120, 250, page_size[0] - inch - 70 - 120 - 250],
        repeatRows=1,
    )
    detail_table.setStyle(_report_table_style())
    story.append(detail_table)

    def draw_footer(canvas, doc) -> None:
        canvas.saveState()
        canvas.setFont("DashboardVera", 8)
        canvas.setFillColor(colors.HexColor("#667085"))
        canvas.drawRightString(
            page_size[0] - 0.5 * inch, 0.25 * inch, f"Page {doc.page}"
        )
        canvas.restoreState()

    document.build(story, onFirstPage=draw_footer, onLaterPages=draw_footer)
    return buffer.getvalue()


def _report_table_style() -> TableStyle:
    return TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dce8f5")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#17365d")),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#c8d6e5")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]
    )


@st.fragment(run_every=REFRESH_SECONDS)
def render_dashboard() -> None:
    st.title("Recruitment Profiles Dashboard")
    st.caption(
        f"Source: shared Google Drive workbook · {SHEET_NAME} sheet · "
        f"Automatically checks for workbook changes every {REFRESH_SECONDS} seconds."
    )

    workbook_url = os.environ.get("SHARED_EXCEL_URL", "").strip()
    if not workbook_url:
        st.error(
            "Configure SHARED_EXCEL_URL in the dashboard host with a Google "
            "Drive sharing link."
        )
        st.stop()

    try:
        file_contents = download_google_drive_workbook(workbook_url)
    except (RuntimeError, URLError, TimeoutError, ValueError) as error:
        st.error(f"Could not load the shared Google Drive workbook: {error}")
        st.stop()

    try:
        profiles = load_profiles(file_contents)
    except (BadZipFile, ValueError) as error:
        st.error(f"Could not read the Google Drive workbook: {error}")
        st.stop()

    with st.sidebar:
        st.header("Filters")
        st.caption("Choose one or more values from each dropdown. Leave a list empty to include all.")
        consultants = sorted(profiles["Consultant"].unique().tolist())
        job_titles = sorted(profiles["Job Title"].unique().tolist())
        statuses = sorted(profiles["Status"].unique().tolist())

        selected_consultants = st.multiselect(
            "Consultant",
            consultants,
            placeholder="All consultants",
            help="Type to search or scroll through the available consultants.",
        )
        selected_job_titles = st.multiselect(
            "Job title",
            job_titles,
            placeholder="All job titles",
            help="Type to search or scroll through the available job titles.",
        )
        selected_statuses = st.multiselect(
            "Status",
            statuses,
            placeholder="All statuses",
            help="Type to search or scroll through the available statuses.",
        )

        valid_dates = profiles["Date"].dropna()
        date_range: tuple[date, date] | date | None = None
        if not valid_dates.empty:
            min_date = valid_dates.min().date()
            max_date = valid_dates.max().date()
            date_range = st.date_input(
                "Profile shared date",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date,
            )
        else:
            st.info("No readable dates were found in the Master Sheet.")

    consultant_filter = selected_consultants or consultants
    job_title_filter = selected_job_titles or job_titles
    status_filter = selected_statuses or statuses
    filtered = profiles[
        profiles["Consultant"].isin(consultant_filter)
        & profiles["Job Title"].isin(job_title_filter)
        & profiles["Status"].isin(status_filter)
    ].copy()

    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
        filtered = filtered[
            filtered["Date"].notna()
            & filtered["Date"].dt.date.ge(start_date)
            & filtered["Date"].dt.date.le(end_date)
        ]
    elif isinstance(date_range, date):
        filtered = filtered[
            filtered["Date"].notna() & filtered["Date"].dt.date.eq(date_range)
        ]

    if isinstance(date_range, tuple) and len(date_range) == 2:
        date_filter_summary = (
            f"{date_range[0].isoformat()} to {date_range[1].isoformat()}"
        )
    elif isinstance(date_range, date):
        date_filter_summary = date_range.isoformat()
    else:
        date_filter_summary = "All dates"
    filter_summary = (
        f"Consultant: {', '.join(selected_consultants) or 'All'}; "
        f"Job title: {', '.join(selected_job_titles) or 'All'}; "
        f"Status: {', '.join(selected_statuses) or 'All'}; "
        f"Date: {date_filter_summary}"
    )
    with st.sidebar:
        st.divider()
        st.subheader("Export report")
        st.download_button(
            "Download PDF report",
            data=build_pdf_report(filtered, filter_summary),
            file_name="recruitment_profiles_report.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    total, by_consultant, by_title, by_status = st.columns(4)
    total.metric("Profiles shared", f"{len(filtered):,}")
    by_consultant.metric("Consultants", f"{filtered['Consultant'].nunique():,}")
    by_title.metric("Job titles", f"{filtered['Job Title'].nunique():,}")
    by_status.metric("Statuses", f"{filtered['Status'].nunique():,}")

    charts = create_dashboard_charts(filtered)
    consultant_column, title_column = st.columns(2)
    with consultant_column:
        st.subheader("Profiles by Consultant")
        st.plotly_chart(
            charts["Profiles by Consultant"],
            use_container_width=True,
        )

    with title_column:
        st.subheader("Profiles by Job Title")
        st.plotly_chart(
            charts["Profiles by Job Title"],
            use_container_width=True,
        )

    st.subheader("Profiles by Consultant and Job Title")
    st.plotly_chart(
        charts["Profiles by Consultant and Job Title"],
        use_container_width=True,
    )

    status_column, date_column = st.columns(2)
    with status_column:
        st.subheader("Profiles by Status")
        st.plotly_chart(
            charts["Profiles by Status"],
            use_container_width=True,
        )

    with date_column:
        st.subheader("Profiles Shared by Date and Consultant")
        st.plotly_chart(
            charts["Profiles Shared by Date and Consultant"],
            use_container_width=True,
        )

    st.subheader("Filtered Profile Summary")
    st.dataframe(
        filtered.loc[:, list(REQUIRED_COLUMNS)]
        .sort_values("Date", ascending=False, na_position="last"),
        use_container_width=True,
        hide_index=True,
    )


render_dashboard()