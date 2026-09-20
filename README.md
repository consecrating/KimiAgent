# KimiAgent

**Turn a topic or an outline into a designer-quality PowerPoint deck — and a matching PDF.**

KimiAgent *reasons a narrative first, then designs the slides*. It emits **native
PowerPoint** — real charts, real tables, real shapes on real slides — and a
**native PDF** drawn from the same design system. Open the `.pptx` in PowerPoint,
Keynote or Google Slides and keep editing; send the `.pdf` to anyone.

> Design inspiration: this project follows the document-generation philosophy of
> [Kimi / Moonshot AI](https://github.com/MoonshotAI) — *native depth, structured
> narrative, professional themes* — but is an independent, open implementation
> built on [`python-pptx`](https://python-pptx.readthedocs.io/) and
> [`reportlab`](https://docs.reportlab.com/). It is not affiliated with or
> endorsed by Moonshot AI.

---

## ✨ Highlights

- **No API key required.** A deterministic offline generator builds a complete,
  coherent deck from just a topic. Zero configuration, no network.
- **PPTX *and* PDF**, both first-class. The PDF is drawn directly — no
  LibreOffice, no Office, no headless converter to install.
- **16 slide archetypes**, including big-statement heroes, KPI cards, feature
  card grids, timelines, process flows and recommendation comparisons.
- **13 themes** with gradients, soft elevation, rounded-corner shape language
  and contrast that is *computed*, not assumed.
- **Text cannot overflow.** Type is measured before it is drawn, and content
  that does not fit flows onto a continuation slide instead of being dropped.
- **Native, editable output** — real charts (column/bar/line/area/pie/doughnut/
  stacked) and real tables, not screenshots.
- **Editable intermediate spec.** Every deck is plain JSON you can export,
  hand-tune and re-render.
- **Speaker notes** on content slides, and a `critique` command that reviews a
  deck's design health.

---

## 🚀 Quick start

```bash
# 1. Install
pip install -r requirements.txt     # python-pptx + reportlab

# 2. Generate a deck AND a PDF from a topic — no API key needed
python -m kimiagent generate \
    --topic "The State of Renewable Energy in 2026" \
    --theme slate \
    --format both \
    -o energy
```

That writes `energy.pptx` and `energy.pdf`, ready to present.

Install as a command (optional):

```bash
pip install -e .
kimiagent generate --topic "Q3 Business Review" --theme corporate -f both -o q3
```

---

## 🧑‍💻 Usage

### Generate from a topic

```bash
python -m kimiagent generate --topic "AI in Healthcare" -o ai.pptx
```

### Both formats, or PDF only

```bash
python -m kimiagent generate -t "AI in Healthcare" --format both -o ai    # ai.pptx + ai.pdf
python -m kimiagent generate -t "AI in Healthcare" --format pdf  -o ai.pdf
```

### Steer the structure with an outline

```bash
python -m kimiagent generate \
    --topic "Q3 Business Review" \
    --outline "Highlights; Revenue; Product; Risks; Next quarter" \
    --audience "Board of directors" \
    --theme corporate \
    --footer-note "Confidential" \
    --format both -o q3
```

Sections you name are always kept. If the slide budget is tight, KimiAgent drops
supporting slides rather than your outline.

### Review a deck's design health

`--report` (or the `critique` command) prints layout statistics and the warnings
a designer would raise on a first pass:

```bash
python -m kimiagent generate -t "Renewables" --report -o r.pptx
```

```
[kimiagent] 19 slides · 67% non-bullet content slides
[kimiagent] layouts: bullets×4, section×4, cards×2, agenda×1, big_statement×1, ...
[kimiagent] design notes: none — deck looks healthy.
```

### Export the spec, edit it, re-render

```bash
python -m kimiagent generate -t "AI in Healthcare" --emit-spec ai.json -o ai.pptx
# ...tweak ai.json by hand, then:
python -m kimiagent from-spec ai.json --format both -o ai
python -m kimiagent critique  ai.json
```

### Tune the look

```bash
python -m kimiagent from-spec ai.json --density dense --aspect 4:3 -o ai.pptx
```

| Option | Values | Effect |
|--------|--------|--------|
| `--theme` | any of the 13 themes | Colour, type and shape language |
| `--density` | `airy` · `normal` · `dense` | Scales the whole type ramp |
| `--aspect` | `16:9` · `4:3` | Canvas size; the grid adapts |
| `--pdf-engine` | `native` · `soffice` · `auto` | How the PDF is produced |

### List themes (and render a sample of each)

```bash
python -m kimiagent themes
python -m kimiagent themes --sample build/samples    # one PDF per theme
```

### Python API

```python
from kimiagent import DeckAgent, render_deck, render_pdf, critique

deck = DeckAgent().build_offline(
    topic="Scaling a Product Team from 5 to 50",
    audience="Engineering leadership",
    theme="aurora",
)
render_deck(deck, "scaling.pptx")
render_pdf(deck, "scaling.pdf")
print(critique(deck))
```

---

## 🔌 Optional: use the Kimi (Moonshot) API for richer content

The offline generator is the default and needs nothing. To let an LLM write the
content instead:

```bash
pip install openai
export KIMI_API_KEY="sk-..."        # or MOONSHOT_API_KEY / OPENAI_API_KEY

python -m kimiagent generate --topic "The Future of Urban Mobility" --llm -f both -o mobility
```

KimiAgent points the OpenAI-compatible client at Moonshot's endpoint by default
(`https://api.moonshot.ai/v1`). Override with `--base-url` / `--model` or the
`KIMI_BASE_URL` / `KIMI_MODEL` environment variables to use any other
OpenAI-compatible provider.

Two safeguards apply to the LLM path: if the API is unavailable KimiAgent
**degrades gracefully** to the offline generator, and if the model returns a deck
missing its cover or closing those are **added back**, because a repaired deck is
more useful than an error.

---

## 🎨 Themes

| Theme | Feel |
|-------|------|
| `kimi` | Deep navy ink + warm coral. Calm, modern *(default)* |
| `slate` | Neutral slate + electric blue, hairline cards. Clean SaaS/product |
| `glacier` | Icy blues on near-white. Airy, analytical, data-heavy |
| `mint` | Fresh teal + mint. Health, climate, wellbeing |
| `sunset` | Warm amber-to-coral gradient on cream. Energetic, human |
| `editorial` | Magazine serif on warm paper, hard edges. Long-form |
| `noir` | Near-black + antique gold, serif headings. Luxury |
| `aurora` | Indigo-to-cyan on deep violet. Dark, vivid, keynote-ready |
| `midnight` | Dark canvas, vivid accents. Tech and product |
| `corporate` | Trustworthy boardroom blue |
| `forest` | Greens for sustainability, energy, nature |
| `morandi` | Muted, dusty, gallery-like tones |
| `inkwash` | Monochrome, editorial, high contrast |

Define your own in a few lines:

```python
from kimiagent.themes import Theme, register_theme

register_theme(Theme(
    name="acme", primary="12203A", secondary="FF5A1F",
    background="FFFFFF", surface="F4F6F9",
    text="111827", text_muted="6B7280", text_on_primary="FFFFFF",
    radius=14, shadow=True,
))
```

Only those fields are required — gradients, borders, chart gridlines and the
dark/light decision are derived for you, and the gradient is automatically
adjusted until one text colour is legible across it.

---

## 🧱 The deck spec

Decks are plain JSON. A minimal example:

```json
{
  "title": "My Deck",
  "theme": "slate",
  "slides": [
    { "type": "cover", "title": "My Deck", "subtitle": "A subtitle" },
    { "type": "big_statement", "kicker": "The point",
      "title": "One claim that deserves a whole slide." },
    { "type": "stats", "title": "Impact", "stats": [
        { "value": "42%", "label": "growth", "detail": "YoY", "trend": "up" }
    ]},
    { "type": "cards", "title": "Drivers", "cards": [
        { "title": "Scale", "body": "Costs fell as volume rose.", "icon": "1" }
    ]},
    { "type": "chart", "title": "Trend", "chart": {
        "chart_type": "line",
        "categories": ["2024", "2025", "2026"],
        "series": [{ "name": "Revenue", "values": [10, 18, 27] }],
        "takeaway": "Growth accelerated after the 2025 launch."
    }},
    { "type": "closing", "title": "Thank you" }
  ]
}
```

See [`examples/sample_spec.json`](examples/sample_spec.json) for a full deck and
[`examples/showcase.py`](examples/showcase.py) for a deck exercising every
archetype in any theme:

```bash
python examples/showcase.py aurora      # -> showcase.pptx + showcase.pdf
```

### Slide types

**Structure** — `cover` · `agenda` · `section` · `closing`
**Prose** — `bullets` · `two_column` · `quote` · `big_statement`
**Structured** — `stats` · `cards` · `timeline` · `process` · `comparison`
**Data** — `chart` · `table` · `image`

Unknown or aliased types (`roadmap`, `pros_cons`, `KPIs`, `feature_cards`, …)
are mapped to the closest real archetype rather than rejected, so lightly
off-schema JSON still renders.

---

## 🏗️ Architecture

```
topic / outline
      │
      ▼
┌───────────────┐   reason narrative,      ┌──────────────┐
│   DeckAgent   │──  choose archetypes  ──▶│  Deck (spec) │  ← plain JSON, editable
│ offline / LLM │                          └──────────────┘
└───────────────┘                                 │
                                 ┌────────────────┴────────────────┐
                                 ▼                                 ▼
                        ┌────────────────┐                ┌────────────────┐
                        │  DeckRenderer  │                │  PDFRenderer   │
                        │  (python-pptx) │                │  (reportlab)   │
                        └────────────────┘                └────────────────┘
                                 │        both driven by            │
                                 │   design.py + layout.py          │
                                 ▼                                 ▼
                             deck.pptx                          deck.pdf
```

| Module | Responsibility |
|--------|----------------|
| `kimiagent/design.py` | The design system: 12-column grid, modular type scale, font metrics, WCAG colour maths |
| `kimiagent/themes.py` | Colour / typography / elevation tokens, contrast-safe gradients |
| `kimiagent/models.py` | The `Deck` / `Slide` spec (JSON round-trip) |
| `kimiagent/layout.py` | Slide zones, card grids, and overflow pagination |
| `kimiagent/agent.py` | Content generation (offline + optional Kimi API) + `critique` |
| `kimiagent/prompts.py` | LLM prompt templates |
| `kimiagent/renderer.py` | Native `.pptx` rendering |
| `kimiagent/pdf_renderer.py` | Native `.pdf` rendering |
| `kimiagent/cli.py` | Command-line interface |

### Why two renderers instead of converting?

Converting `.pptx` to PDF needs LibreOffice or Office installed, which fails on
most CI runners and containers, and re-flows text through a different layout
engine so the PDF quietly stops matching the deck. Drawing the PDF directly from
the same grid, type scale and pagination keeps the two in agreement and keeps
install requirements to two pip packages.

If you *do* have LibreOffice and want byte-level fidelity to the `.pptx`, ask for
it: `--pdf-engine soffice` (or `auto` to prefer it when present).

### How overflow is prevented

`python-pptx`'s `fit_text` needs a font file on disk and only supports a 12 pt
baseline, and PowerPoint's own autofit is computed by the *renderer* — nothing is
actually resized until a human opens the file. KimiAgent instead measures text
up front with embedded Helvetica/Arial metrics (exact to the PDF backend's own
font), so it can:

1. shrink type to fit its box before drawing,
2. avoid splitting a word mid-way when a column is narrow,
3. and flow genuinely over-full content onto a continuation slide.

The test suite renders every archetype in every theme, then re-opens the saved
file and re-measures every text box to assert nothing overflows.

---

## 🧪 Development

```bash
pip install -e ".[dev]"
pytest -q
```

## 📦 Requirements

- Python 3.9+
- [`python-pptx`](https://python-pptx.readthedocs.io/) — `.pptx` output
- [`reportlab`](https://docs.reportlab.com/) — `.pdf` output
- [`openai`](https://pypi.org/project/openai/) *(optional — LLM path only)*
- [`pillow`](https://pypi.org/project/pillow/) *(optional — aspect-correct image placement)*

## 📄 License

MIT — see [LICENSE](LICENSE).
