"""Small, reusable UI fragments for the Retail Sales Performance app.

The iframe markup is authored here and contains no user-generated HTML. Any
dynamic text inserted into the iframe is HTML-escaped before rendering.
"""

from __future__ import annotations

from html import escape
import textwrap

import streamlit as st


TAILWIND_CDN = "https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4.3.5"


def _render_html_fragment(markup: str, height: int, alt: str) -> None:
    """Use the supported iframe API, with the prior Streamlit HTML component as fallback."""
    if hasattr(st, "iframe"):
        st.iframe(markup, width="stretch", height=height, tab_index=0, alt=alt)
    else:
        st.components.v1.html(markup, height=height, scrolling=False)


def _page(fragment: str) -> str:
    """Wrap static component markup in a responsive HTML document."""
    return textwrap.dedent(
        f"""\
        <!doctype html>
        <html lang="en">
          <head>
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <script src="{TAILWIND_CDN}"></script>
            <style>
              html, body {{ margin: 0; padding: 0; background: transparent; }}
              body {{ font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }}
            </style>
          </head>
          <body class="bg-transparent text-slate-900">
            {fragment}
          </body>
        </html>
        """
    )


def render_header(title: str = "Retail Sales Performance Analysis", subtitle: str = "Synthetic data | Jan 2023 – Dec 2024") -> None:
    """Render an accessible project header; title and subtitle are escaped."""
    safe_title = escape(str(title))
    safe_subtitle = escape(str(subtitle))
    markup = f"""
    <header class="flex flex-col gap-3 rounded-2xl bg-slate-950 px-5 py-4 text-white sm:flex-row sm:items-center sm:justify-between">
      <div>
        <p class="text-xs font-semibold uppercase tracking-[0.18em] text-teal-300">ABES · MCA Data Analytics</p>
        <h1 class="mt-1 text-2xl font-bold leading-tight">{safe_title}</h1>
        <p class="mt-1 text-sm text-slate-300">{safe_subtitle}</p>
      </div>
      <span class="inline-flex w-fit items-center rounded-full border border-teal-300/40 bg-teal-400/10 px-3 py-1 text-xs font-semibold text-teal-200">SYNTHETIC DATA</span>
    </header>
    """
    _render_html_fragment(_page(markup), 116, "Project title, dataset period, and synthetic-data notice")


def render_kpi_card(
    label: str,
    value: str | int | float,
    delta: str | None = None,
    delta_kind: str = "neutral",
    help_text: str | None = None,
) -> None:
    """Render one accessible KPI card using escaped display values.

    delta_kind accepts only: positive, negative, or neutral. These values select
    fixed Tailwind classes; user or dataset values are never used as class names.
    """
    tones = {
        "positive": "text-emerald-700 bg-emerald-50",
        "negative": "text-rose-700 bg-rose-50",
        "neutral": "text-slate-600 bg-slate-100",
    }
    tone = tones.get(delta_kind, tones["neutral"])
    safe_label = escape(str(label))
    safe_value = escape(str(value))
    safe_delta = escape(str(delta)) if delta is not None else ""
    safe_help = escape(str(help_text)) if help_text else ""
    delta_markup = (
        f'<span class="mt-2 inline-flex rounded-full px-2.5 py-1 text-xs font-semibold {tone}">{safe_delta}</span>'
        if delta is not None else ""
    )
    help_markup = f'<p class="mt-2 text-xs text-slate-500">{safe_help}</p>' if safe_help else ""
    markup = f"""
    <article class="rounded-2xl border border-slate-200 bg-white px-5 py-4 shadow-sm">
      <p class="text-xs font-semibold uppercase tracking-wide text-slate-500">{safe_label}</p>
      <p class="mt-2 break-words text-3xl font-bold tracking-tight text-slate-950">{safe_value}</p>
      {delta_markup}
      {help_markup}
    </article>
    """
    _render_html_fragment(_page(markup), 142 if help_text else 116, f"Key performance indicator: {safe_label}, {safe_value}")


def render_empty_state(message: str, detail: str | None = None) -> None:
    """Show a readable native empty state without inventing zero metrics."""
    st.info(message)
    if detail:
        st.caption(detail)


def render_insight(message: str, limitation: str | None = None) -> None:
    """Render a concise insight with an optional interpretation boundary."""
    st.success(message)
    if limitation:
        st.caption(limitation)
