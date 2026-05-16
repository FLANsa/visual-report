"""Visual Report — لوحة تحكم تقرير القنوات."""

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
        direction: rtl;
        text-align: right;
        font-family: 'Segoe UI', 'Tahoma', 'Geneva', 'Verdana', sans-serif;
      }
      [data-testid="stMetric"] { text-align: right; }
      [data-testid="stMetricLabel"] { justify-content: flex-end; }
      [data-testid="stMetricValue"] { direction: ltr; text-align: right; }
      .stPlotlyChart { direction: ltr; }
      h1, h2, h3, h4 { font-weight: 700; }
      .stDataFrame { direction: ltr; }
      .block-container { padding-top: 2rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


def _seconds_from_time(value) -> float:
    """تحويل الوقت لثواني للحسابات."""
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
    """عرض الثواني بصيغة HH:MM:SS."""
    if pd.isna(total_seconds):
        return "00:00:00"
    total_seconds = int(round(total_seconds))
    h, rem = divmod(total_seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


@st.cache_data(show_spinner=False)
def load_channels_report(file_bytes: bytes, source_name: str) -> pd.DataFrame:
    """قراءة ومعالجة ملف تقرير القنوات."""
    df = pd.read_excel(
        BytesIO(file_bytes),
        sheet_name="Channels Report",
        header=2,
        usecols="G:O",
        engine="openpyxl",
    )
    df.columns = [str(c).strip() for c in df.columns]

    df["Date"] = df["Date"].ffill()

    # دمج الأعمدة المنقسمة للـ Pending (بسبب الدمج K:L في الإكسل)
    extra_col = next((c for c in df.columns if c.lower().startswith("unnamed")), None)
    if extra_col:
        df["Pending"] = df["Pending"].fillna(0) + df[extra_col].fillna(0)
        df = df.drop(columns=[extra_col])

    df = df.dropna(subset=["Channel"]).copy()
    df["Channel"] = df["Channel"].astype(str).str.strip()

    for col in ["Incoming", "Closed", "Pending", "Backlog"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    df["Service Level"] = pd.to_numeric(df["Service Level"], errors="coerce").fillna(0)

    df["Response Seconds"] = df["Average Response"].apply(_seconds_from_time)
    df["Average Response Text"] = df["Response Seconds"].apply(_format_seconds)

    df["Closure Rate"] = (df["Closed"] / df["Incoming"].where(df["Incoming"] > 0, pd.NA)).fillna(0)

    df = df.reset_index(drop=True)
    df.attrs["source"] = source_name
    return df


def render_header() -> None:
    st.title("📊 Visual Report")
    st.markdown(
        "ارفع ملف Excel الخاص بتقرير القنوات وستظهر لك لوحة تحكم تفاعلية "
        "بأهم المؤشرات والرسوم البيانية لأداء القنوات."
    )


def get_data_source() -> tuple[bytes, str] | None:
    """رجّع (file_bytes, source_name) أو None لو ما رفع شيء."""
    uploaded = st.file_uploader(
        "📁 ارفع ملف Excel",
        type=["xlsx", "xls"],
        help="يجب أن يحتوي الملف على ورقة باسم 'Channels Report' بنفس الهيكل.",
    )

    if uploaded is not None:
        return uploaded.getvalue(), uploaded.name
    return None


def render_welcome() -> None:
    st.info(
        "👋 مرحباً! ارفع ملف Excel من الأعلى لعرض لوحة التحكم."
    )
    with st.expander("📋 ما هو هيكل الملف المطلوب؟"):
        st.markdown(
            """
            - **اسم الورقة:** Channels Report
            - **الأعمدة:** Date, Channel, Incoming, Closed, Pending, Backlog, Average Response, Service Level
            - **يدعم:** الخلايا المدموجة في عمود التاريخ، وقيم Pending المنقسمة على عمودين
            """
        )


def render_kpis(df: pd.DataFrame) -> None:
    total_backlog = int(df["Backlog"].sum())

    st.metric(
        "📦 إجمالي المتراكم",
        f"{total_backlog:,}",
        delta_color="inverse" if total_backlog > 0 else "normal",
    )


def _plotly_layout(title: str, **extra) -> dict:
    base = dict(
        title=dict(text=title, x=0.98, xanchor="right", font=dict(size=16)),
        font=dict(family="Tahoma, Segoe UI, sans-serif", size=12),
        margin=dict(t=60, b=40, l=20, r=20),
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    base.update(extra)
    return base


def chart_incoming_vs_closed(df: pd.DataFrame) -> go.Figure:
    sorted_df = df.sort_values("Incoming", ascending=False)
    fig = go.Figure()
    fig.add_bar(
        name="الوارد",
        x=sorted_df["Channel"],
        y=sorted_df["Incoming"],
        marker_color=PRIMARY,
        text=sorted_df["Incoming"],
        textposition="outside",
    )
    fig.add_bar(
        name="المُغلق",
        x=sorted_df["Channel"],
        y=sorted_df["Closed"],
        marker_color=SUCCESS,
        text=sorted_df["Closed"],
        textposition="outside",
    )
    fig.update_layout(
        **_plotly_layout("الوارد مقابل المُغلق لكل قناة"),
        barmode="group",
        xaxis=dict(tickangle=-35),
        yaxis=dict(title="عدد التذاكر", gridcolor="#eee"),
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
        name="مستوى الخدمة",
    )
    fig.add_vline(
        x=SLA_THRESHOLD * 100,
        line_dash="dash",
        line_color=NEUTRAL,
        annotation_text=f"الهدف {SLA_THRESHOLD:.0%}",
        annotation_position="top",
    )
    fig.update_layout(
        **_plotly_layout("مستوى الخدمة لكل قناة", showlegend=False),
        xaxis=dict(title="مستوى الخدمة (٪)", range=[0, 110], gridcolor="#eee"),
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
        name="متوسط الاستجابة",
    )
    fig.update_layout(
        **_plotly_layout("متوسط وقت الاستجابة (دقيقة)", showlegend=False),
        xaxis=dict(tickangle=-35),
        yaxis=dict(title="الدقائق", gridcolor="#eee"),
    )
    return fig


def chart_pending_backlog(df: pd.DataFrame) -> go.Figure:
    sorted_df = df.sort_values(["Backlog", "Pending"], ascending=False)
    fig = go.Figure()
    fig.add_bar(
        name="المُعلّق",
        x=sorted_df["Channel"],
        y=sorted_df["Pending"],
        marker_color=WARNING,
    )
    fig.add_bar(
        name="المتراكم",
        x=sorted_df["Channel"],
        y=sorted_df["Backlog"],
        marker_color=DANGER,
    )
    fig.update_layout(
        **_plotly_layout("المُعلّق والمتراكم لكل قناة"),
        barmode="stack",
        xaxis=dict(tickangle=-35),
        yaxis=dict(title="عدد التذاكر", gridcolor="#eee"),
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
    fig.update_layout(**_plotly_layout("توزيع الوارد على القنوات"))
    return fig


def render_charts(df: pd.DataFrame) -> None:
    st.subheader("📈 الرسوم البيانية")
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
            f"### 📅 تاريخ التقرير: {pd.Timestamp(report_date).strftime('%Y-%m-%d')}"
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
            "❌ تعذّر قراءة الملف. تأكد أن الملف يحتوي على ورقة باسم "
            "'Channels Report' بنفس الهيكل المتوقع."
        )
        st.exception(exc)
        return

    if df.empty:
        st.warning("الملف لا يحتوي على بيانات قنوات صالحة.")
        return

    render_dashboard(df)


if __name__ == "__main__":
    main()
