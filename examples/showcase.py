"""Build a deck that exercises every slide archetype, in any theme.

Useful both as a reference for the spec format and as a quick way to see what
a theme actually looks like across all sixteen layouts.

    python examples/showcase.py                  # kimi theme -> showcase.pptx + .pdf
    python examples/showcase.py aurora           # a different theme
    python examples/showcase.py slate out/deck   # custom output stem
"""

from __future__ import annotations

import sys

from kimiagent.models import (
    Bullet,
    Card,
    ChartData,
    ChartSeries,
    ComparisonSide,
    Deck,
    ProcessStep,
    Slide,
    SlideType,
    Stat,
    TableData,
    TimelineItem,
)


def build_showcase(theme: str = "kimi", density: str = "normal", aspect: str = "16:9") -> Deck:
    """A realistic 19-slide deck touching all sixteen archetypes."""
    return Deck(
        title="The State of Renewable Energy",
        subtitle="Grid parity, storage economics and the next decade",
        author="Research Desk",
        theme=theme,
        density=density,
        aspect=aspect,
        footer_note="Internal — Confidential",
        slides=[
            Slide(
                type=SlideType.COVER,
                title="The State of Renewable Energy in 2026",
                subtitle="Why the cost curve just became a policy problem",
                kicker="Annual outlook",
            ),
            Slide(
                type=SlideType.AGENDA,
                title="What we'll cover",
                bullets=[
                    Bullet(text="Where costs actually landed"),
                    Bullet(text="Storage: the real bottleneck"),
                    Bullet(text="Grid interconnection queues"),
                    Bullet(text="Policy levers that worked"),
                    Bullet(text="What to watch in 2027"),
                ],
            ),
            Slide(type=SlideType.SECTION, title="Where costs actually landed", subtitle="01"),
            Slide(
                type=SlideType.BIG_STATEMENT,
                title="Solar is now the cheapest electricity in history.",
                kicker="The headline",
                subtitle="And it is still getting cheaper.",
            ),
            Slide(
                type=SlideType.STATS,
                title="The numbers that moved",
                kicker="At a glance",
                subtitle="Global weighted averages, 2020 to 2026",
                stats=[
                    Stat(value="-71%", label="Utility solar LCOE", detail="since 2015", trend="down"),
                    Stat(value="$28", label="Per MWh, best sites", detail="unsubsidised", trend="down"),
                    Stat(value="4.2x", label="Storage deployed", detail="vs. 2021", trend="up"),
                    Stat(value="38%", label="Of new capacity", detail="renewables share", trend="up"),
                ],
            ),
            Slide(
                type=SlideType.CHART,
                title="Cost per MWh keeps falling",
                kicker="Evidence",
                chart=ChartData(
                    chart_type="line",
                    unit="USD/MWh",
                    categories=["2020", "2021", "2022", "2023", "2024", "2025", "2026"],
                    series=[
                        ChartSeries("Utility solar", [52, 46, 41, 38, 34, 31, 28]),
                        ChartSeries("Onshore wind", [48, 45, 43, 40, 38, 36, 35]),
                        ChartSeries("Gas peaker", [74, 79, 96, 88, 84, 86, 89]),
                    ],
                    takeaway="Solar crossed below gas in 2021 and the gap has widened every year since.",
                ),
                notes="Lead with the crossover point, not the absolute numbers.",
            ),
            Slide(
                type=SlideType.CARDS,
                title="Four forces behind the curve",
                cards=[
                    Card(title="Manufacturing scale", body="Polysilicon capacity tripled, collapsing module prices.", icon="1"),
                    Card(title="Cheaper capital", body="Projects now price like infrastructure, not venture bets.", icon="2"),
                    Card(title="Better siting", body="Higher capacity factors from smarter placement.", icon="3"),
                    Card(title="Policy certainty", body="Long-dated credits let developers plan a decade out.", icon="4"),
                ],
            ),
            Slide(type=SlideType.SECTION, title="Storage is the real bottleneck", subtitle="02"),
            Slide(
                type=SlideType.TIMELINE,
                title="How storage economics unlocked",
                subtitle="Cost per kWh at pack level",
                timeline=[
                    TimelineItem(period="2019", title="$180/kWh", detail="Pilot scale only", done=True),
                    TimelineItem(period="2022", title="$130/kWh", detail="First merchant projects", done=True),
                    TimelineItem(period="2024", title="$95/kWh", detail="4-hour duration standard", done=True),
                    TimelineItem(period="2026", title="$71/kWh", detail="8-hour becomes bankable"),
                    TimelineItem(period="2029", title="<$50/kWh", detail="Seasonal shifting viable"),
                ],
            ),
            Slide(
                type=SlideType.PROCESS,
                title="How a project reaches the grid",
                steps=[
                    ProcessStep(title="Site & secure", detail="Land, resource data, offtake letter."),
                    ProcessStep(title="Interconnect", detail="Queue position and impact studies."),
                    ProcessStep(title="Finance", detail="Debt sizing against the PPA."),
                    ProcessStep(title="Build", detail="Procurement and construction."),
                    ProcessStep(title="Operate", detail="Dispatch against price signals."),
                ],
            ),
            Slide(
                type=SlideType.COMPARISON,
                title="Where to put the next dollar",
                verdict="Fund storage. The generation cost curve no longer needs our help.",
                sides=[
                    ComparisonSide(
                        title="More generation",
                        subtitle="Add panels and turbines",
                        points=[
                            Bullet(text="Proven, fast to deploy"),
                            Bullet(text="Costs already falling without us"),
                            Bullet(text="Curtailment rising in sunny markets", bold_lead="Risk"),
                            Bullet(text="Marginal value per MW is declining"),
                        ],
                    ),
                    ComparisonSide(
                        title="More storage",
                        subtitle="Shift supply in time",
                        preferred=True,
                        points=[
                            Bullet(text="Directly attacks the binding constraint"),
                            Bullet(text="Captures the evening price spread"),
                            Bullet(text="Defers costly transmission upgrades"),
                            Bullet(text="Supply chain is tightening", bold_lead="Risk"),
                        ],
                    ),
                ],
            ),
            Slide(
                type=SlideType.TABLE,
                title="Interconnection queues by region",
                kicker="The bottleneck, quantified",
                table=TableData(
                    headers=["Region", "Queued GW", "Avg wait", "Completion rate"],
                    numeric_columns=[1, 3],
                    highlight_row=4,
                    rows=[
                        ["PJM", "286", "4.3 yrs", "14%"],
                        ["MISO", "241", "5.1 yrs", "11%"],
                        ["ERCOT", "198", "1.8 yrs", "31%"],
                        ["CAISO", "154", "3.6 yrs", "19%"],
                        ["Total / avg", "879", "3.7 yrs", "18%"],
                    ],
                ),
            ),
            Slide(
                type=SlideType.CHART,
                title="Where the queued capacity sits",
                chart=ChartData(
                    chart_type="doughnut",
                    data_labels=True,
                    categories=["Solar", "Storage", "Wind", "Gas", "Other"],
                    series=[ChartSeries("Share", [46, 28, 17, 6, 3])],
                    takeaway="Three quarters of the queue is solar and storage.",
                ),
            ),
            Slide(
                type=SlideType.BULLETS,
                title="What actually unblocks a queue",
                subtitle="Lessons from the fastest-moving markets",
                bullets=[
                    Bullet(text="Charge for study work so speculative applications self-select out.", bold_lead="Price the queue"),
                    Bullet(text="Cluster studies instead of serial ones"),
                    Bullet(text="ERCOT processes in batches and clears 3x faster", level=1),
                    Bullet(text="Publish hosting capacity maps before applications open.", bold_lead="Transparency"),
                    Bullet(text="Let projects pay for their own upgrades and recover later.", bold_lead="Cost allocation"),
                ],
                notes="This is the operational heart of the deck.",
            ),
            Slide(
                type=SlideType.TWO_COLUMN,
                title="Two readings of the same data",
                columns=[
                    [
                        Bullet(text="The optimist", bold_lead="View"),
                        Bullet(text="Costs fall, deployment compounds"),
                        Bullet(text="Storage solves intermittency by 2030"),
                    ],
                    [
                        Bullet(text="The realist", bold_lead="View"),
                        Bullet(text="Physical grid build-out is the limit"),
                        Bullet(text="Permitting timelines dominate everything"),
                    ],
                ],
            ),
            Slide(
                type=SlideType.QUOTE,
                quote="We stopped having a generation problem and started having a delivery problem.",
                attribution="Grid operator, PJM interconnection review",
            ),
            Slide(
                type=SlideType.IMAGE,
                title="Reference architecture",
                caption="Figure 1 — Co-located solar and storage plant topology",
            ),
            Slide(
                type=SlideType.CLOSING,
                title="Thank you",
                subtitle="Questions, and the data room is open",
            ),
        ],
    )


def main(argv: list[str]) -> int:
    from kimiagent.pdf_renderer import render_pdf
    from kimiagent.renderer import render_deck

    theme = argv[1] if len(argv) > 1 else "kimi"
    stem = argv[2] if len(argv) > 2 else "showcase"
    deck = build_showcase(theme)
    print("Wrote", render_deck(deck, f"{stem}.pptx"))
    print("Wrote", render_pdf(deck, f"{stem}.pdf"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
