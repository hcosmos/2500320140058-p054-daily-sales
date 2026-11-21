"""
PROJECT P054: DAILY SALES ACTUAL VS TARGET ANALYSIS
UI Component Library & Tailwind CSS Framework Module (src/ui.py)
Author: Harshit (University Roll No: 2500320140058) | Project ID: P054 | Batch ID: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 | Project Mentor: Dinesh Sereno
"""

import streamlit as st
import streamlit.components.v1 as components

TAILWIND_CDN = '<script src="https://cdn.tailwindcss.com"></script>'

def render_tailwind_header(title: str, subtitle: str, meta_tag: str = "Enterprise Retail Performance Analytics"):
    """
    Renders an executive, responsive Tailwind CSS header banner.
    """
    html_content = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        {TAILWIND_CDN}
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
        <style>
            body {{ font-family: 'Inter', sans-serif; margin: 0; padding: 0; background-color: transparent; }}
        </style>
    </head>
    <body class="p-1">
        <div class="bg-gradient-to-r from-blue-900 via-indigo-900 to-slate-900 rounded-xl p-5 shadow-lg border border-slate-800 text-white">
            <div class="flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div>
                    <div class="inline-flex items-center gap-2 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-500/20 text-blue-300 border border-blue-400/30 mb-2">
                        <span class="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse"></span>
                        {meta_tag}
                    </div>
                    <h1 class="text-2xl md:text-3xl font-bold tracking-tight text-white mb-1">{title}</h1>
                    <p class="text-slate-300 text-sm max-w-2xl">{subtitle}</p>
                </div>
                <div class="hidden lg:flex flex-col items-end text-xs text-slate-400 border-l border-slate-700/60 pl-4 py-1">
                    <span class="font-medium text-slate-200">Release 1.0 (MVP)</span>
                    <span>Candidate: Harshit (ABES)</span>
                    <span>Mentor: Dinesh Sereno</span>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html_content, height=140, scrolling=False)


def render_kpi_cards(kpi_data: list):
    """
    Renders a grid of responsive Tailwind CSS KPI cards.
    kpi_data: list of dicts with keys:
      - 'label': str
      - 'value': str (formatted currency/percent)
      - 'delta': str
      - 'is_positive': bool
      - 'subtext': str
    """
    cards_html = ""
    for card in kpi_data:
        delta_color = "text-emerald-600 bg-emerald-50 border-emerald-200" if card.get("is_positive", True) else "text-rose-600 bg-rose-50 border-rose-200"
        arrow = "↑" if card.get("is_positive", True) else "↓"
        cards_html += f"""
        <div class="bg-white rounded-xl p-4 shadow-sm border border-slate-200 hover:shadow-md transition-shadow">
            <div class="text-xs font-medium text-slate-500 uppercase tracking-wider mb-1">{card.get('label', '')}</div>
            <div class="text-xl md:text-2xl font-bold text-slate-800 tracking-tight mb-2">{card.get('value', '')}</div>
            <div class="flex items-center justify-between">
                <span class="inline-flex items-center px-2 py-0.5 rounded text-xs font-semibold border {delta_color}">
                    {arrow} {card.get('delta', '')}
                </span>
                <span class="text-xs text-slate-400">{card.get('subtext', '')}</span>
            </div>
        </div>
        """

    full_html = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        {TAILWIND_CDN}
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
        <style>
            body {{ font-family: 'Inter', sans-serif; margin: 0; padding: 2px; background-color: transparent; }}
        </style>
    </head>
    <body>
        <div class="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-3">
            {cards_html}
        </div>
    </body>
    </html>
    """
    components.html(full_html, height=130, scrolling=False)


def render_insight_box(title: str, message: str, alert_type: str = "info"):
    """
    Renders an alert/callout box with Tailwind styling.
    alert_type: 'success', 'warning', 'danger', 'info'
    """
    themes = {
        "success": ("bg-emerald-50 border-emerald-300 text-emerald-900", "text-emerald-700", "✓"),
        "warning": ("bg-amber-50 border-amber-300 text-amber-900", "text-amber-700", "⚠"),
        "danger": ("bg-rose-50 border-rose-300 text-rose-900", "text-rose-700", "✕"),
        "info": ("bg-blue-50 border-blue-300 text-blue-900", "text-blue-700", "ℹ")
    }
    box_class, text_class, icon = themes.get(alert_type, themes["info"])

    html_content = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        {TAILWIND_CDN}
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
        <style>
            body {{ font-family: 'Inter', sans-serif; margin: 0; padding: 2px; background-color: transparent; }}
        </style>
    </head>
    <body>
        <div class="rounded-lg border p-3.5 {box_class} shadow-sm">
            <div class="flex items-start gap-3">
                <span class="text-base font-bold {text_class}">{icon}</span>
                <div>
                    <h4 class="text-xs font-bold uppercase tracking-wider {text_class} mb-0.5">{title}</h4>
                    <p class="text-xs leading-relaxed">{message}</p>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html_content, height=80, scrolling=False)
