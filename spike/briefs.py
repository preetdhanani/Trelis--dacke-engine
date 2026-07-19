"""S2 -- ~20 varied briefs per the spike method (short, long, messy, multi-point,
comparison, data-bearing), spanning all 8 manifest layouts. None of these name a
layout_id explicitly -- picking the right one is part of what's being tested.
"""

BRIEFS = [
    # 1. short -- title
    "Open the deck with our standard title slide for the Acme Corp cost-reduction engagement.",
    # 2. short -- section divider
    "Kick off section 2 on Talent Strategy.",
    # 3. long narrative -- text_bullets
    (
        "We need a slide explaining why now is the right moment for Acme to enter the "
        "direct-to-consumer channel. Retail margins have compressed for three straight "
        "years, their two largest wholesale partners have signaled they're cutting shelf "
        "space in the next planning cycle, and a DTC pilot they ran informally last year "
        "actually outperformed wholesale unit economics by a wide margin even without any "
        "real investment behind it. This should land as a clear argument slide."
    ),
    # 4. messy/typos -- text_bullets
    (
        "ok so basically the client keeps asking abt margins going down we shd probably "
        "jsut put a bullet slide together sayin why - competition, input costs up, "
        "discounting too much idk maybe 4-5 bullets"
    ),
    # 5. multi-point, user-supplied list -- text_bullets
    (
        "Summarize these five findings as bullets: 1) Revenue flat year over year "
        "2) Costs up 9% driven by logistics 3) Churn improved two points "
        "4) NPS held steady at 61 5) New product launch delayed to Q2"
    ),
    # 6. comparison -- two_column
    (
        "Compare our proposed pricing tiers against the two main competitors. Narrative on "
        "the left explaining why our tiering is simpler and easier to sell, we'll drop in a "
        "comparison graphic on the right."
    ),
    # 7. data-bearing -- exhibit_data
    (
        "Show quarterly revenue for the last four quarters: Q1 $2.1M, Q2 $2.4M, Q3 $2.9M, "
        "Q4 $3.4M, and make the point that growth is accelerating each quarter."
    ),
    # 8. quote/testimonial -- quote_callout
    (
        "We got a great quote from the client CFO after the workshop: 'This engagement paid "
        "for itself in the first month.' Put that on its own slide, attribute it to the CFO."
    ),
    # 9. closing/contact -- closing_contact
    "Wrap up the deck with contact info: Jane Doe, Senior Partner, jane@firm.com, 555-2020.",
    # 10. image-only
    (
        "We need a slide that's basically just the org chart photo we took of the whiteboard "
        "session, with a small caption noting it's the current-state org."
    ),
    # 11. data-bearing exhibit with takeaway + source
    (
        "Build an exhibit slide on hiring: we hired 40 people this year vs a target of 25. "
        "Chart on the left, a one-line takeaway on the right about how this outpaced plan, "
        "cite HR data as the source."
    ),
    # 12. long narrative, two_column with sizing logic
    (
        "For the market sizing slide: TAM is $4.2B growing at 11% CAGR, our addressable "
        "slice is roughly $600M given the segments we can realistically serve in year one. "
        "We want a supporting market-map image on the right, with 4-5 bullets on the left "
        "explaining the sizing logic and why the addressable slice is defensible."
    ),
    # 13. vague/ambiguous -- tests graceful degradation
    "something about our competitive advantage",
    # 14. messy comparison, no chart wanted -- text_bullets
    (
        "so basically we beat competitor X on speed n price but they beat us on brand — "
        "need a slide that shows this, prob bullets is fine no chart needed"
    ),
    # 15. messy section divider
    "next section is gonna be about Risks, just want a clean divider, number it 03.",
    # 16. very short data point -- quote_callout as a stat callout
    "One number: NPS is 62, up from 48 last year. Make it a big callout, quote-style slide.",
    # 17. long messy multi-point exec summary -- text_bullets
    (
        "ok for the exec summary bullets we want: strong Q3 performance, but watch margin "
        "trend closely, new market opens Q1 next year, one competitor exited the space which "
        "helps us, and there's a legal thing pending we should mention briefly, keep it to "
        "like 5-6 bullets max nothing too dramatic on the legal bit"
    ),
    # 18. data-bearing time series -- exhibit_data
    (
        "Plot the cost-per-acquisition trend over the last 6 months: 42, 39, 37, 35, 33, 30 "
        "(dollars) and make the point CAC is trending down consistently. Exhibit slide, "
        "source noted as internal ads dashboard."
    ),
    # 19. two-column narrative with a diagram-ish visual
    (
        "Explain the new onboarding flow -- write out the 4 steps as bullets and we'll put "
        "the flow diagram image on the right side of the slide."
    ),
    # 20. quote hybrid for close of pitch
    (
        "Take this client testimonial and turn it into the punchy quote slide for the close "
        "of the pitch: 'Your team moved faster than any vendor we've worked with, and the "
        "results showed up in the numbers within 60 days.' Attribute to VP Operations, "
        "client name withheld."
    ),
]

assert len(BRIEFS) == 20
