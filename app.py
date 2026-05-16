"""Visual Report — channels dashboard."""

from __future__ import annotations

import datetime as dt
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

PRIMARY = "#1f77b4"
SUCCESS = "#2ca02c"
WARNING = "#ff9f1c"
DANGER = "#d62728"
NEUTRAL = "#7f8c8d"
SLA_THRESHOLD = 0.80

st.set_page_config(
    page_title="Visual Report",
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
      html, body, [class*="css"], .main, .block-container {
        direction: ltr;
        text-align: left;
        font-family: 'Segoe UI', 'Tahoma', 'Geneva', 'Verdana', sans-serif;
      }
      [data-testid="stMetric"] { text-align: left; }
      [data-testid="stMetricLabel"] { justify-content: flex-start; }
      [data-testid="stMetricValue"] { direction: ltr; text-align: left; }
      .stPlotlyChart { direction: ltr; }
      h1, h2, h3, h4 { font-weight: 700; }
      .stDataFrame { direction: ltr; }
      .block-container { padding-top: 2rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


def _seconds_from_time(value) -> float:
    """Convert a time value to seconds."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, dt.time):
        return value.hour * 3600 + value.minute * 60 + value.second
    if isinstance(value, dt.timedelta):
        return value.total_seconds()
    if isinstance(value, (int, float)):
        return float(value) * 86400 if value < 1 else float(value)
    text = str(value).strip()
    try:
        parts = [int(p) for p in text.split(":")]
    except ValueError:
        return 0.0
    while len(parts) < 3:
        parts.insert(0, 0)
    h, m, s = parts[-3:]
    return h * 3600 + m * 60 + s


def _format_seconds(total_seconds: float) -> str:
    """Format seconds as HH:MM:SS."""
    if pd.isna(total_seconds):
        return "00:00:00"
    total_seconds = int(round(total_seconds))
    h, rem = divmod(total_seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename = {
        "AverageResponse": "Average Response",
        "ServiceLevel": "Service Level",
    }
    df = df.rename(columns={c: rename.get(str(c).strip(), str(c).strip()) for c in df.columns})
    return df


@st.cache_data(show_spinner=False)
def load_channels_report(file_bytes: bytes, source_name: str) -> pd.DataFrame:
    """Load and parse a channels report workbook."""
    buffer = BytesIO(file_bytes)
    workbook = pd.ExcelFile(buffer, engine="openpyxl")

    if "Channels Report" in workbook.sheet_names:
        df = pd.read_excel(
            workbook,
            sheet_name="Channels Report",
            header=2,
            usecols="G:O",
        )
    elif "Data" in workbook.sheet_names:
        df = pd.read_excel(workbook, sheet_name="Data", header=0)
    else:
        raise ValueError(
            "Could not find a 'Channels Report' or 'Data' sheet. "
            f"Available sheets: {', '.join(workbook.sheet_names)}"
        )

    df = _normalize_columns(df)
    df["Date"] = df["Date"].ffill()

    extra_col = next((c for c in df.columns if c.lower().startswith("unnamed")), None)
    if extra_col:
        df["Pending"] = df["Pending"].fillna(0) + df[extra_col].fillna(0)
        df = df.drop(columns=[extra_col])

    df = df.dropna(subset=["Channel"]).copy()
    df["Channel"] = df["Channel"].astype(str).str.strip()

    for col in ["Incoming", "Closed", "Pending", "Backlog"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    service = df["Service Level"] if "Service Level" in df.columns else df.get("ServiceLevel")
    df["Service Level"] = pd.to_numeric(service, errors="coerce").fillna(0)
    df.loc[df["Service Level"] > 1, "Service Level"] /= 100

    if "ResponseSeconds" in df.columns:
        df["Response Seconds"] = pd.to_numeric(df["ResponseSeconds"], errors="coerce").fillna(0)
    else:
        response_col = "Average Response" if "Average Response" in df.columns else "AverageResponse"
        df["Response Seconds"] = df[response_col].apply(_seconds_from_time)
    df["Average Response Text"] = df["Response Seconds"].apply(_format_seconds)

    df["Closure Rate"] = (df["Closed"] / df["Incoming"].where(df["Incoming"] > 0, pd.NA)).fillna(0)

    df = df.reset_index(drop=True)
    df.attrs["source"] = source_name
    return df


def render_header() -> None:
    st.title("📊 Visual Report")
    st.markdown(
        "Upload your channels Excel report to view an interactive dashboard "
        "with key metrics and charts for channel performance."
    )


def get_data_source() -> tuple[bytes, str] | None:
    """Return (file_bytes, source_name) or None if nothing was uploaded."""
    uploaded = st.file_uploader(
        "📁 Upload Excel file",
        type=["xlsx", "xls"],
        help="The file should include a 'Channels Report' or 'Data' sheet with the expected structure.",
    )

    if uploaded is not None:
        return uploaded.getvalue(), uploaded.name
    return None


def render_welcome() -> None:
    st.info(
        "👋 Welcome! Upload an Excel file above to view the dashboard."
    )
    with st.expander("📋 Required file structure"):
        st.markdown(
            """
            - **Sheet name:** Channels Report (or Data)
            - **Columns:** Date, Channel, Incoming, Closed, Pending, Backlog, Average Response, Service Level
            - **Supports:** merged date cells and Pending values split across two columns
            """
        )


def render_kpis(df: pd.DataFrame) -> None:
    total_backlog = int(df["Backlog"].sum())

    st.metric(
        "📦 Total backlog",
        f"{total_backlog:,}",
        delta_color="inverse" if total_backlog > 0 else "normal",
    )


def _plotly_layout(title: str, **extra) -> dict:
    base = dict(
        title=dict(text=title, x=0.02, xanchor="left", font=dict(size=16)),
        font=dict(family="Tahoma, Segoe UI, sans-serif", size=12),
        margin=dict(t=60, b=40, l=20, r=20),
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    )
    base.update(extra)
    return base


def chart_incoming_vs_closed(df: pd.DataFrame) -> go.Figure:
    sorted_df = df.sort_values("Incoming", ascending=False)
    fig = go.Figure()
    fig.add_bar(
        name="Incoming",
        x=sorted_df["Channel"],
        y=sorted_df["Incoming"],
        marker_color=PRIMARY,
        text=sorted_df["Incoming"],
        textposition="outside",
    )
    fig.add_bar(
        name="Closed",
        x=sorted_df["Channel"],
        y=sorted_df["Closed"],
        marker_color=SUCCESS,
        text=sorted_df["Closed"],
        textposition="outside",
    )
    fig.update_layout(
        **_plotly_layout("Incoming vs closed by channel"),
        barmode="group",
        xaxis=dict(tickangle=-35),
        yaxis=dict(title="Ticket count", gridcolor="#eee"),
    )
    return fig


def chart_service_level(df: pd.DataFrame) -> go.Figure:
    sorted_df = df.sort_values("Service Level", ascending=True)
    colors = [SUCCESS if v >= SLA_THRESHOLD else DANGER for v in sorted_df["Service Level"]]
    fig = go.Figure()
    fig.add_bar(
        x=sorted_df["Service Level"] * 100,
        y=sorted_df["Channel"],
        orientation="h",
        marker_color=colors,
        text=[f"{v:.0%}" for v in sorted_df["Service Level"]],
        textposition="outside",
        name="Service level",
    )
    fig.add_vline(
        x=SLA_THRESHOLD * 100,
        line_dash="dash",
        line_color=NEUTRAL,
        annotation_text=f"Target {SLA_THRESHOLD:.0%}",
        annotation_position="top",
    )
    fig.update_layout(
        **_plotly_layout("Service level by channel", showlegend=False),
        xaxis=dict(title="Service level (%)", range=[0, 110], gridcolor="#eee"),
        yaxis=dict(title=""),
    )
    return fig


def chart_response_time(df: pd.DataFrame) -> go.Figure:
    sorted_df = df.sort_values("Response Seconds", ascending=False)
    fig = go.Figure()
    fig.add_bar(
        x=sorted_df["Channel"],
        y=sorted_df["Response Seconds"] / 60,
        marker_color=WARNING,
        text=sorted_df["Average Response Text"],
        textposition="outside",
        name="Avg response",
    )
    fig.update_layout(
        **_plotly_layout("Average response time (minutes)", showlegend=False),
        xaxis=dict(tickangle=-35),
        yaxis=dict(title="Minutes", gridcolor="#eee"),
    )
    return fig


def chart_pending_backlog(df: pd.DataFrame) -> go.Figure:
    sorted_df = df.sort_values(["Backlog", "Pending"], ascending=False)
    fig = go.Figure()
    fig.add_bar(
        name="Pending",
        x=sorted_df["Channel"],
        y=sorted_df["Pending"],
        marker_color=WARNING,
    )
    fig.add_bar(
        name="Backlog",
        x=sorted_df["Channel"],
        y=sorted_df["Backlog"],
        marker_color=DANGER,
    )
    fig.update_layout(
        **_plotly_layout("Pending and backlog by channel"),
        barmode="stack",
        xaxis=dict(tickangle=-35),
        yaxis=dict(title="Ticket count", gridcolor="#eee"),
    )
    return fig


def chart_volume_pie(df: pd.DataFrame) -> go.Figure:
    pie_df = df[df["Incoming"] > 0].copy()
    fig = px.pie(
        pie_df,
        values="Incoming",
        names="Channel",
        hole=0.45,
    )
    fig.update_traces(textposition="inside", textinfo="percent+label")
    fig.update_layout(**_plotly_layout("Incoming volume by channel"))
    return fig


def render_charts(df: pd.DataFrame) -> None:
    st.subheader("📈 Charts")
    row1_col1, row1_col2 = st.columns(2)
    with row1_col1:
        st.plotly_chart(chart_incoming_vs_closed(df), use_container_width=True)
    with row1_col2:
        st.plotly_chart(chart_service_level(df), use_container_width=True)

    row2_col1, row2_col2 = st.columns(2)
    with row2_col1:
        st.plotly_chart(chart_response_time(df), use_container_width=True)
    with row2_col2:
        st.plotly_chart(chart_pending_backlog(df), use_container_width=True)

    st.plotly_chart(chart_volume_pie(df), use_container_width=True)


def render_dashboard(df: pd.DataFrame) -> None:
    report_date = df["Date"].dropna().iloc[0] if df["Date"].notna().any() else None
    if report_date is not None:
        st.markdown(
            f"### 📅 Report date: {pd.Timestamp(report_date).strftime('%Y-%m-%d')}"
        )

    st.divider()
    render_kpis(df)
    st.divider()
    render_charts(df)


def main() -> None:
    render_header()
    source = get_data_source()

    if source is None:
        render_welcome()
        return

    file_bytes, source_name = source
    try:
        df = load_channels_report(file_bytes, source_name)
    except Exception as exc:
        st.error(
            "❌ Could not read the file. Make sure it includes a 'Channels Report' or "
            "'Data' sheet with the expected structure."
        )
        st.exception(exc)
        return

    if df.empty:
        st.warning("The file does not contain valid channel data.")
        return

    render_dashboard(df)


if __name__ == "__main__":
    main()
