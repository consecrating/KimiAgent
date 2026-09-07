"""Prompt templates for the LLM-backed content generation path.

Kept in one place so the "reason the narrative first" instructions can be
tuned without touching the agent logic.
"""

from __future__ import annotations

SYSTEM_PROMPT = """You are KimiAgent, an expert presentation designer and \
narrative strategist. You turn a topic or raw material into a tight, \
persuasive slide deck.

Work in two stages, internally:
1. Reason the argument into shape: decide the story arc, the key message of \
each section, and what evidence supports it.
2. Design slides that carry that argument, choosing the most fitting slide \
type for each idea.

Rules:
- Prefer clarity and momentum over cramming. 3-6 bullets per slide, each a \
short, punchy phrase (not a paragraph).
- Use a variety of slide types. Use `stats` for quantified impact, `chart` \
for trends/comparisons (invent realistic illustrative numbers when exact data \
is unknown, and note that they are illustrative), `two_column` for \
comparisons/pros-cons, `quote` for a memorable idea, `table` for structured \
comparisons.
- Every deck starts with a `cover`, has an `agenda`, uses `section` dividers \
between major parts, and ends with a `closing`.
- Write speaker notes for content slides.

Return ONLY valid JSON matching the provided schema. No markdown, no commentary."""


def build_user_prompt(topic: str, slide_count: int, audience: str, tone: str, extra: str) -> str:
    parts = [
        f"Create a presentation deck about: {topic}",
        f"Target length: about {slide_count} slides (including cover, agenda, sections and closing).",
    ]
    if audience:
        parts.append(f"Audience: {audience}")
    if tone:
        parts.append(f"Tone: {tone}")
    if extra:
        parts.append(f"Additional context / source material:\n{extra}")
    parts.append(SCHEMA_HINT)
    return "\n".join(parts)


SCHEMA_HINT = """
Return JSON with this exact shape:

{
  "title": "string",
  "subtitle": "string",
  "author": "string",
  "theme": "kimi | morandi | inkwash | corporate | forest | midnight",
  "slides": [
    {"type": "cover", "title": "...", "subtitle": "..."},
    {"type": "agenda", "title": "Agenda", "bullets": ["Item 1", "Item 2"]},
    {"type": "section", "title": "Section name", "subtitle": "optional"},
    {"type": "bullets", "title": "...", "bullets": [
        {"text": "point", "level": 0, "bold_lead": "optional lead-in"}
    ], "notes": "speaker notes"},
    {"type": "two_column", "title": "...", "columns": [["a","b"], ["c","d"]]},
    {"type": "stats", "title": "...", "stats": [
        {"value": "42%", "label": "growth", "detail": "optional"}
    ]},
    {"type": "chart", "title": "...", "chart": {
        "chart_type": "column | bar | line | pie | area",
        "categories": ["2022","2023","2024"],
        "series": [{"name": "Series A", "values": [1,2,3]}],
        "unit": "optional (e.g. %, $B)"
    }, "notes": "..."},
    {"type": "table", "title": "...", "table": {
        "headers": ["Col A","Col B"],
        "rows": [["1","2"],["3","4"]]
    }},
    {"type": "quote", "quote": "...", "attribution": "..."},
    {"type": "closing", "title": "Thank you", "subtitle": "call to action"}
  ]
}

Bullets may be plain strings or objects. Keep it valid JSON.
"""
