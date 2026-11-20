import streamlit as st
import streamlit.components.v1 as components

def render_tailwind_header(title="Treatment Cost Trend Analysis", subtitle="Executive Health Economics & Billing Intelligence Dashboard"):
    """Renders an executive header using Tailwind CSS CDN via st.components.v1.html."""
    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <script src="https://cdn.tailwindcss.com"></script>
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
            body {{ font-family: 'Inter', sans-serif; margin: 0; padding: 0; background-color: transparent; }}
        </style>
    </head>
    <body class="p-2">
        <div class="bg-gradient-to-r from-[#1B365D] via-[#2C3E50] to-[#A51C30] rounded-xl p-5 text-white shadow-lg border border-slate-700">
            <div class="flex flex-col md:flex-row justify-between items-start md:items-center">
                <div>
                    <div class="flex items-center space-x-2">
                        <span class="bg-[#A51C30] text-xs font-bold px-2.5 py-1 rounded-full uppercase tracking-wider text-white">MCA Project 033</span>
                        <span class="text-xs text-slate-300 font-medium">ABES Engineering College</span>
                    </div>
                    <h1 class="text-2xl md:text-3xl font-extrabold tracking-tight mt-1">{title}</h1>
                    <p class="text-xs md:text-sm text-slate-200 mt-1">{subtitle}</p>
                </div>
                <div class="mt-3 md:mt-0 text-right">
                    <p class="text-xs text-slate-300 font-semibold">Author: Arohi Tyagi (2500320140033)</p>
                    <p class="text-xs text-slate-300">Mentor: Prof. Alpana Lodhi</p>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=130)

def render_kpi_card(title, value, delta=None, delta_positive=True, note=None):
    """Renders a single high-impact KPI card styled with Tailwind CSS."""
    delta_color = "text-emerald-600 bg-emerald-50" if delta_positive else "text-rose-600 bg-rose-50"
    delta_icon = "↑" if delta_positive else "↓"
    delta_html = f'<span class="text-xs font-bold px-2 py-0.5 rounded {delta_color} ml-2">{delta_icon} {delta}</span>' if delta else ''
    note_html = f'<p class="text-xs text-slate-400 mt-1">{note}</p>' if note else ''

    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <script src="https://cdn.tailwindcss.com"></script>
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
            body {{ font-family: 'Inter', sans-serif; margin: 0; padding: 0; background-color: transparent; }}
        </style>
    </head>
    <body class="p-1">
        <div class="bg-white rounded-xl p-4 border border-slate-200 shadow-sm hover:shadow-md transition duration-200">
            <p class="text-xs font-semibold uppercase tracking-wider text-slate-500">{title}</p>
            <div class="flex items-baseline mt-2">
                <span class="text-2xl font-bold text-[#1B365D]">{value}</span>
                {delta_html}
            </div>
            {note_html}
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=95)

def render_kpi_row(kpis):
    """Renders 4 KPI cards side by side in Streamlit."""
    cols = st.columns(len(kpis))
    for idx, kpi in enumerate(kpis):
        with cols[idx]:
            render_kpi_card(
                title=kpi.get("title", ""),
                value=kpi.get("value", ""),
                delta=kpi.get("delta"),
                delta_positive=kpi.get("delta_positive", True),
                note=kpi.get("note")
            )

def render_insight_box(title, text, style="info"):
    """Renders an analytical insight callout box."""
    colors = {
        "info": ("bg-blue-50", "border-blue-300", "text-blue-900", "Info"),
        "warning": ("bg-amber-50", "border-amber-300", "text-amber-900", "Alert"),
        "success": ("bg-emerald-50", "border-emerald-300", "text-emerald-900", "Success"),
        "danger": ("bg-rose-50", "border-rose-300", "text-rose-900", "Outlier Warning")
    }
    bg, border, text_col, badge = colors.get(style, colors["info"])

    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="p-1">
        <div class="{bg} {border} border-l-4 rounded-r-lg p-3 shadow-xs">
            <div class="flex items-center space-x-2">
                <span class="text-xs font-bold uppercase tracking-wider px-2 py-0.5 rounded {bg} {text_col} border">{badge}</span>
                <span class="text-sm font-semibold {text_col}">{title}</span>
            </div>
            <p class="text-xs {text_col} mt-1.5 leading-relaxed">{text}</p>
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=85)
