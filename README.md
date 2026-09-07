# KimiAgent

**Turn a topic or an outline into a native, designer-quality PowerPoint deck.**

KimiAgent is a small, self-contained agent that *reasons a narrative first, then
designs the slides*. It emits **native PowerPoint** — real charts, real tables,
real shapes on real slides — rather than dumping text into a template or handing
you flat text boxes. Open the output in PowerPoint, Keynote or Google Slides and
keep editing: the charts and tables are fully editable.

> Design inspiration: this project follows the document-generation philosophy of
> [Kimi / Moonshot AI](https://github.com/MoonshotAI) — *native depth, structured
> narrative, professional themes* — but is an independent, open implementation
> built on [`python-pptx`](https://python-pptx.readthedocs.io/). It is not
> affiliated with or endorsed by Moonshot AI.

---

## ✨ Highlights

- **No API key required.** A deterministic offline generator builds a complete,
  coherent deck from just a topic. Zero configuration, no network.
- **Optional LLM path.** If you *do* have a Kimi / Moonshot (or any
  OpenAI-compatible) API key, KimiAgent will use it to write richer content.
- **Native, editable output** — native charts (column/bar/line/pie/area) and
  native tables, not screenshots.
- **11 slide archetypes** — cover, agenda, section dividers, bullets, two-column,
  KPI stats, chart, table, quote, image, closing.
- **6 professional themes** — `kimi`, `morandi`, `inkwash`, `corporate`,
  `forest`, `midnight`.
- **Editable intermediate spec.** Every deck is a plain JSON document you can
  export, hand-tune and re-render.
- **Speaker notes** generated for content slides.

---

## 🚀 Quick start

```bash
# 1. Install the one dependency
pip install python-pptx     # or: pip install -r requirements.txt

# 2. Generate a deck from a topic — no API key needed
python -m kimiagent generate \
    --topic "The State of Renewable Energy in 2026" \
    --theme forest \
    -o energy.pptx
```

That writes a ready-to-present `energy.pptx`. Open it in any slides app.

Install as a command (optional):

```bash
pip install -e .
kimiagent generate --topic "Q3 Business Review" --theme corporate -o q3.pptx
```

---

## 🧑‍💻 Usage

### Generate from a topic

```bash
python -m kimiagent generate --topic "AI in Healthcare" -o ai.pptx
```

### Steer the structure with an outline

```bash
python -m kimiagent generate \
    --topic "Q3 Business Review" \
    --outline "Highlights; Revenue; Product; Risks; Next quarter" \
    --audience "Board of directors" \
    --theme corporate \
    -o q3.pptx
```

### Export the spec, edit it, re-render

```bash
# Generate a deck AND save its editable JSON spec
python -m kimiagent generate --topic "AI in Healthcare" --emit-spec ai.json -o ai.pptx

# ...tweak ai.json by hand, then:
python -m kimiagent from-spec ai.json -o ai.pptx
```

### List themes

```bash
python -m kimiagent themes
```

### Python API

```python
from kimiagent import DeckAgent, render_deck

deck = DeckAgent().build_offline(
    topic="Scaling a Product Team from 5 to 50",
    audience="Engineering leadership",
    theme="kimi",
)
render_deck(deck, "scaling.pptx")
```

---

## 🔌 Optional: use the Kimi (Moonshot) API for richer content

The offline generator is the default and needs nothing. To let an LLM write the
content instead, install the SDK and set a key:

```bash
pip install openai
export KIMI_API_KEY="sk-..."        # or MOONSHOT_API_KEY / OPENAI_API_KEY

python -m kimiagent generate --topic "The Future of Urban Mobility" --llm -o mobility.pptx
```

KimiAgent points the OpenAI-compatible client at Moonshot's endpoint by default
(`https://api.moonshot.ai/v1`). Override with `--base-url` / `--model` or the
`KIMI_BASE_URL` / `KIMI_MODEL` environment variables to use any other
OpenAI-compatible provider. If the API is unavailable, KimiAgent **degrades
gracefully** to the offline generator so you always get a deck.

---

## 🎨 Themes

| Theme | Feel |
|-------|------|
| `kimi` | Deep navy ink + warm coral accent — calm, modern (default) |
| `morandi` | Muted, dusty gallery tones |
| `inkwash` | Monochrome, editorial, high-contrast |
| `corporate` | Trustworthy boardroom blue |
| `forest` | Greens for sustainability / energy / nature |
| `midnight` | Dark background with vivid accents — tech/product |

---

## 🧱 The deck spec

Decks are plain JSON. A minimal example:

```json
{
  "title": "My Deck",
  "theme": "kimi",
  "slides": [
    { "type": "cover", "title": "My Deck", "subtitle": "A subtitle" },
    { "type": "bullets", "title": "Key points", "bullets": ["First", "Second"] },
    { "type": "stats", "title": "Impact", "stats": [
        { "value": "42%", "label": "growth", "detail": "YoY" }
    ]},
    { "type": "chart", "title": "Trend", "chart": {
        "chart_type": "line",
        "categories": ["2024", "2025", "2026"],
        "series": [{ "name": "Revenue", "values": [10, 18, 27] }]
    }},
    { "type": "closing", "title": "Thank you" }
  ]
}
```

See [`examples/sample_spec.json`](examples/sample_spec.json) for a full deck, and
[`examples/generate_example.py`](examples/generate_example.py) for the Python API.

### Slide types

`cover` · `agenda` · `section` · `bullets` · `two_column` · `stats` · `chart` ·
`table` · `quote` · `image` · `closing`

---

## 🏗️ Architecture

```
topic / outline
      │
      ▼
┌───────────────┐   reason narrative,      ┌──────────────┐
│   DeckAgent   │──  choose slide types ──▶│  Deck (spec) │  ← plain JSON, editable
│ offline / LLM │                          └──────────────┘
└───────────────┘                                 │
                                                   ▼
                                          ┌────────────────┐
                                          │  DeckRenderer  │  native pptx shapes,
                                          │  (python-pptx) │  charts & tables
                                          └────────────────┘
                                                   │
                                                   ▼
                                              deck.pptx
```

| Module | Responsibility |
|--------|----------------|
| `kimiagent/models.py` | The `Deck` / `Slide` data model (JSON round-trip) |
| `kimiagent/themes.py` | Colour palettes + typography |
| `kimiagent/agent.py` | Content generation (offline + optional Kimi API) |
| `kimiagent/prompts.py` | LLM prompt templates |
| `kimiagent/renderer.py` | Native `.pptx` rendering |
| `kimiagent/cli.py` | Command-line interface |

---

## 📦 Requirements

- Python 3.9+
- [`python-pptx`](https://python-pptx.readthedocs.io/) (required)
- [`openai`](https://pypi.org/project/openai/) (optional — only for the LLM path)

## 📄 License

MIT — see [LICENSE](LICENSE).
