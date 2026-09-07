"""Build a professional credentials + portfolio deck for Sanctify.

Sanctify is an award-winning advertising & digital marketing agency in Vasco,
South Goa, India (https://www.sanctify.in). This script assembles a branded
deck that opens with the agency story and services, then showcases the graphic
design portfolio (19 pieces sourced from the agency's design document) as a
gallery.

Run from the repository root::

    python examples/sanctify/build_sanctify_deck.py

Output: examples/sanctify/Sanctify-Digital-Marketing-Goa.pptx
No API key required — this is a hand-authored deck spec rendered by KimiAgent.
"""

import os
import sys

# Make the package importable when run directly from the repo (no install).
HERE = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, REPO_ROOT)

from kimiagent import (
    Deck,
    Slide,
    SlideType,
    Bullet,
    Stat,
    ImageItem,
    render_deck,
)

ASSETS = os.path.join(HERE, "assets")


def img(name: str, caption: str = "") -> ImageItem:
    return ImageItem(path=os.path.join(ASSETS, name), caption=caption)


def build(theme: str = "midnight") -> Deck:
    slides = []

    # ---- 1. Cover -------------------------------------------------------
    slides.append(
        Slide(
            type=SlideType.COVER,
            title="SANCTIFY",
            subtitle="Award-winning Advertising & Digital Marketing Agency in Goa",
        )
    )

    # ---- 2. Agenda ------------------------------------------------------
    slides.append(
        Slide(
            type=SlideType.AGENDA,
            title="What's inside",
            bullets=[
                Bullet(text="Who we are"),
                Bullet(text="What we do"),
                Bullet(text="Why Sanctify — results & reach"),
                Bullet(text="Graphic design portfolio"),
                Bullet(text="What our clients say"),
                Bullet(text="Let's work together"),
            ],
        )
    )

    # ---- 3. About -------------------------------------------------------
    slides.append(
        Slide(
            type=SlideType.BULLETS,
            title="Who we are",
            subtitle="Vasco, South Goa · India",
            bullets=[
                Bullet(text="An award-winning advertising & digital marketing agency serving Goa for over 13 years.", bold_lead="Established"),
                Bullet(text="A full suite of services under one roof — strategy, creative, web and performance.", bold_lead="One partner"),
                Bullet(text="A competitive response rate and quantifiable results set us apart.", bold_lead="Measurable"),
                Bullet(text="An 'Amaze Creation Agency' — creative ideas paired with strategic execution.", bold_lead="Creative"),
            ],
            notes="Open with the Sanctify story: 13+ years, Vasco/South Goa, award-winning, full-service, results-driven.",
        )
    )

    # ---- 4. Services (section + two column) -----------------------------
    slides.append(Slide(type=SlideType.SECTION, title="What we do", subtitle="Services"))
    slides.append(
        Slide(
            type=SlideType.TWO_COLUMN,
            title="A full-suite agency",
            columns=[
                [
                    Bullet(text="Digital Marketing", bold_lead="Performance"),
                    Bullet(text="Search Engine Optimization (SEO)"),
                    Bullet(text="Search Engine Marketing (SEM) / PPC"),
                    Bullet(text="Social Media Marketing (SMM)"),
                    Bullet(text="Google, Facebook, Instagram & LinkedIn Ads"),
                    Bullet(text="Email marketing & Bulk SMS"),
                ],
                [
                    Bullet(text="Creative & Web", bold_lead="Brand"),
                    Bullet(text="Website design & development"),
                    Bullet(text="Logo, branding & identity"),
                    Bullet(text="Graphic design — flyers, banners, catalogues"),
                    Bullet(text="Online classifieds"),
                    Bullet(text="Newspaper & outdoor advertising"),
                ],
            ],
            notes="We cover the full journey from brand identity through website to demand generation.",
        )
    )

    # ---- 5. Why Sanctify — stats ---------------------------------------
    slides.append(
        Slide(
            type=SlideType.STATS,
            title="Why Sanctify",
            subtitle="Track record",
            stats=[
                Stat(value="4.8/5", label="Top-rated agency", detail="128 reviews"),
                Stat(value="13+", label="Years of experience", detail="serving Goa"),
                Stat(value="2600+", label="Daily website visitors", detail="sanctify.in"),
                Stat(value="100+", label="Websites developed", detail="and counting"),
            ],
            notes="Credibility markers pulled from the website. Numbers are the agency's stated figures.",
        )
    )
    slides.append(
        Slide(
            type=SlideType.STATS,
            title="Search visibility we deliver",
            subtitle="Organic reach",
            stats=[
                Stat(value="1200+", label="URLs indexed", detail="across the site"),
                Stat(value="800+", label="URLs on Google top pages", detail="for target keywords"),
                Stat(value="1000+", label="Keywords on page 1", detail="of Google.co.in"),
            ],
        )
    )

    # ---- 6. Portfolio section ------------------------------------------
    slides.append(
        Slide(
            type=SlideType.SECTION,
            title="Graphic Design Portfolio",
            subtitle="Selected work",
        )
    )

    # Social & search creatives
    slides.append(
        Slide(
            type=SlideType.GALLERY,
            title="Social & search creatives",
            images=[
                img("image1.png", "Google Map post — thumbnail update"),
                img("image2.png", "Social media post — Hindustan Petroleum (HP Gas)"),
            ],
        )
    )

    # A4 flyer — single portrait
    slides.append(
        Slide(
            type=SlideType.IMAGE,
            title="Print advertising",
            image_path=img("image3.png").path,
            caption="A4 flyer design — Karni Chemicals",
        )
    )

    # Packaging — popcorn tub covers (3 flavours)
    slides.append(
        Slide(
            type=SlideType.GALLERY,
            title="Packaging design — popcorn tub covers",
            subtitle="Flavour range",
            images=[
                img("image5.png", "Chilli Cheese"),
                img("image4.png", "Chilli Tomato"),
                img("image7.png", "Cream Onion"),
            ],
        )
    )

    # Festival campaigns
    slides.append(
        Slide(
            type=SlideType.GALLERY,
            title="Festival campaigns",
            images=[
                img("image6.png", "Road Carrier — festival poster"),
                img("image9.png", "Festival post"),
            ],
        )
    )

    # Catalogue design
    slides.append(
        Slide(
            type=SlideType.GALLERY,
            title="Catalogue design",
            images=[
                img("image8.png", "Inside page"),
                img("image12.png", "Outside page"),
            ],
        )
    )

    # Cinders Biryani flyer
    slides.append(
        Slide(
            type=SlideType.GALLERY,
            title="Cinders Biryani — flyer",
            subtitle="Front & back",
            images=[
                img("image10.png", "Front"),
                img("image11.png", "Back"),
            ],
        )
    )

    # Political / public figure social media
    slides.append(
        Slide(
            type=SlideType.GALLERY,
            title="Social media — Transport Minister Shri. Mauvin Godinho",
            images=[
                img("image13.png"),
                img("image14.png"),
                img("image15.png"),
            ],
        )
    )

    # Restaurant menu design (4-up)
    slides.append(
        Slide(
            type=SlideType.GALLERY,
            title="Restaurant menu design",
            images=[
                img("image16.png"),
                img("image17.png"),
                img("image18.png"),
                img("image19.png"),
            ],
        )
    )

    # ---- 7. Testimonials ------------------------------------------------
    slides.append(Slide(type=SlideType.SECTION, title="What our clients say", subtitle="Testimonials"))
    slides.append(
        Slide(
            type=SlideType.QUOTE,
            quote="Their expertise in SEO, social media and online strategies helped my business rank higher and attract more customers. Professional, responsive, and they truly understand the local market.",
            attribution="Benz Nx — SEO & SMM services",
        )
    )
    slides.append(
        Slide(
            type=SlideType.QUOTE,
            quote="We got our website built by them and continue with their social media marketing — we're very happy with the results. The team responds to our queries with great patience.",
            attribution="Pawan Raj — Web & digital marketing",
        )
    )

    # ---- 8. Why choose us ----------------------------------------------
    slides.append(
        Slide(
            type=SlideType.BULLETS,
            title="Why businesses in Goa choose Sanctify",
            bullets=[
                Bullet(text="Over a decade of local expertise in the Goa market", bold_lead="Local"),
                Bullet(text="Transparent, measurable results you can track", bold_lead="Transparent"),
                Bullet(text="SEO, web design and PPC under one roof", bold_lead="Full-suite"),
                Bullet(text="A creative, results-driven approach focused on ROI", bold_lead="ROI-first"),
            ],
        )
    )

    # ---- 9. Closing / contact ------------------------------------------
    slides.append(
        Slide(
            type=SlideType.CLOSING,
            title="Let's grow your brand",
            subtitle="www.sanctify.in  ·  Advertising & Digital Marketing Agency, Goa",
        )
    )

    return Deck(
        title="Sanctify — Digital Marketing Agency in Goa",
        subtitle="Credentials & Portfolio",
        author="www.sanctify.in",
        theme=theme,
        slides=slides,
    )


def main() -> None:
    # Build two variants so you can pick the look you prefer:
    #   * midnight  — premium dark; colourful creative work pops on dark cards
    #   * corporate — clean, light, boardroom-friendly
    variants = {
        "midnight": "Sanctify-Digital-Marketing-Goa.pptx",
        "corporate": "Sanctify-Digital-Marketing-Goa-Light.pptx",
    }
    theme = sys.argv[1] if len(sys.argv) > 1 else None
    if theme:
        variants = {theme: f"Sanctify-Digital-Marketing-Goa-{theme}.pptx"}

    for theme_name, filename in variants.items():
        deck = build(theme=theme_name)
        out = os.path.join(HERE, filename)
        render_deck(deck, out)
        print(f"[{theme_name}] Built {len(deck.slides)} slides -> {out}")


if __name__ == "__main__":
    main()
