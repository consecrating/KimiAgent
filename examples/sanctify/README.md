# Sanctify — Credentials & Portfolio Deck

A professional 21-slide deck for **Sanctify**, an award-winning advertising &
digital marketing agency in Vasco, South Goa, India ([sanctify.in](https://www.sanctify.in)),
built with KimiAgent.

## What's here

| File | Description |
|------|-------------|
| `Sanctify-Digital-Marketing-Goa.pptx` | **Premium dark** theme (`midnight`) — colourful creative work pops on dark cards |
| `Sanctify-Digital-Marketing-Goa-Light.pptx` | **Clean light** theme (`corporate`) — boardroom-friendly |
| `build_sanctify_deck.py` | The deck builder script |
| `assets/` | The 19 graphic-design portfolio images |

Both `.pptx` files are ready to open in PowerPoint, Keynote or Google Slides.

## Deck flow

1. Cover
2. Agenda
3. Who we are
4. What we do (section) + full-suite services
5. Why Sanctify — track record + search visibility (stats)
6. **Graphic design portfolio** (section) — 8 gallery slides showcasing all 19 pieces:
   social & search creatives, print flyer, popcorn packaging (3 flavours),
   festival campaigns, catalogue, Cinders Biryani flyer, public-figure social
   media, and restaurant menu design
7. Testimonials
8. Why choose Sanctify
9. Contact / closing

## Rebuild

```bash
# From the repository root
python examples/sanctify/build_sanctify_deck.py

# Or a single theme (any of: kimi, morandi, inkwash, corporate, forest, midnight)
python examples/sanctify/build_sanctify_deck.py forest
```

## Content sources

- Agency description, services, stats and testimonials: [sanctify.in](https://www.sanctify.in)
- Portfolio images: the agency's *Graphic Designs* document.
