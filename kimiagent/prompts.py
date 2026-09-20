"""Prompt templates for the LLM-backed content generation path.

Kept in one place so the "reason the narrative first" instructions can be
tuned without touching the agent logic.

The prompt does two jobs. The obvious one is describing the JSON schema. The
less obvious and more important one is pushing the model *away from bullet
lists*: left to itself, a model will answer "make me a deck" with a title and
five bullets per slide, which is exactly the output that reads as machine-made.
The instructions below name the archetypes, say when each is the right choice,
and ask for a takeaway on every chart — so the structure carries the argument
rather than the prose having to.
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

How to choose a slide type — this matters more than the wording:
- `big_statement` for the thesis and for any single claim that deserves the \
whole slide. Use it at least once, early.
- `cards` for three to six parallel, peer-level ideas (drivers, features, \
principles). Prefer this over a bullet list whenever the items are siblings \
rather than a sequence of thoughts.
- `stats` for quantified impact: three or four numbers, each with a short \
label. Add "trend": "up" | "down" | "flat" where there is a direction.
- `chart` for trends and comparisons. ALWAYS include a "takeaway": one \
sentence stating what the chart shows. Invent realistic illustrative numbers \
when exact data is unknown and say so in the takeaway or notes.
- `timeline` for anything that happens in an order over time (roadmaps, \
history, phases).
- `process` for a repeatable sequence of steps (three to five).
- `comparison` for a decision between two options. Mark the better one with \
"preferred": true and state a one-line "verdict".
- `table` for structured multi-attribute comparison.
- `two_column` for exactly two contrasting viewpoints.
- `quote` for a memorable line.
- `bullets` only when the content genuinely is a short list of related points. \
Aim for at most a third of your content slides to be `bullets`.

Rules:
- Prefer clarity and momentum over cramming. Three to six bullets per slide, \
each a short punchy phrase, never a paragraph. Keep bullets under 120 \
characters.
- Every deck starts with a `cover`, has an `agenda`, uses `section` dividers \
between major parts (give each a "subtitle" of "01", "02", ... which renders \
as a large section numeral), and ends with a `closing`.
- Use "kicker" for a short eyebrow label above a title (e.g. "Evidence", \
"At a glance") — it gives the audience the frame before the detail.
- Write speaker notes for the content slides you would actually talk to.

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
Return JSON with this exact shape. Every slide needs "type"; include only the
fields relevant to that type.

{
  "title": "string",
  "subtitle": "string",
  "author": "string",
  "theme": "kimi | slate | aurora | glacier | mint | sunset | editorial | noir | corporate | forest | midnight | morandi | inkwash",
  "footer_note": "optional, e.g. Confidential",
  "slides": [
    {"type": "cover", "title": "...", "subtitle": "...", "kicker": "optional eyebrow"},

    {"type": "agenda", "title": "What we'll cover", "bullets": ["Item 1", "Item 2"]},

    {"type": "section", "title": "Section name", "subtitle": "01"},

    {"type": "big_statement", "kicker": "The argument",
     "title": "One sentence that deserves a whole slide.",
     "subtitle": "optional supporting line"},

    {"type": "bullets", "title": "...", "kicker": "optional", "subtitle": "optional",
     "bullets": [{"text": "point", "level": 0, "bold_lead": "optional lead-in"}],
     "notes": "speaker notes"},

    {"type": "cards", "title": "...", "cards": [
        {"title": "Card title", "body": "One or two short sentences.", "icon": "1"}
    ]},

    {"type": "stats", "title": "...", "kicker": "At a glance", "stats": [
        {"value": "42%", "label": "growth", "detail": "optional", "trend": "up"}
    ]},

    {"type": "timeline", "title": "...", "timeline": [
        {"period": "Q1 2026", "title": "Milestone", "detail": "optional", "done": true}
    ]},

    {"type": "process", "title": "...", "steps": [
        {"title": "Step name", "detail": "optional short line"}
    ]},

    {"type": "comparison", "title": "...", "verdict": "The recommendation, in one line.",
     "sides": [
        {"title": "Option A", "subtitle": "optional", "points": ["short point"]},
        {"title": "Option B", "subtitle": "optional", "preferred": true,
         "points": [{"text": "short point", "bold_lead": "Risk"}]}
     ]},

    {"type": "two_column", "title": "...", "columns": [["a","b"], ["c","d"]]},

    {"type": "chart", "title": "...", "chart": {
        "chart_type": "column | bar | line | area | pie | doughnut | stacked_column | stacked_bar",
        "categories": ["2024","2025","2026"],
        "series": [{"name": "Series A", "values": [1,2,3]}],
        "unit": "optional (e.g. %, $B)",
        "data_labels": false,
        "takeaway": "REQUIRED: one sentence saying what this chart shows."
    }, "notes": "..."},

    {"type": "table", "title": "...", "table": {
        "headers": ["Col A","Col B"],
        "rows": [["1","2"],["3","4"]],
        "numeric_columns": [1],
        "highlight_row": 3
    }},

    {"type": "quote", "quote": "...", "attribution": "..."},

    {"type": "image", "title": "...", "image_path": "path/to.png", "caption": "..."},

    {"type": "closing", "title": "Thank you", "subtitle": "call to action"}
  ]
}

Bullets and card/step/point entries may be plain strings or objects. Keep it
valid JSON.
"""

__all__ = ["SYSTEM_PROMPT", "SCHEMA_HINT", "build_user_prompt"]
