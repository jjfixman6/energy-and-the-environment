"""
SMR (NuScale) Presentation Generator
Generates a professional 16:9 PowerPoint compatible with Keynote, PowerPoint, and Slides.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# --- Color Palette ---
BG_DARK = RGBColor(0x1B, 0x1F, 0x3B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GREY = RGBColor(0xD0, 0xD4, 0xE4)
TEAL = RGBColor(0x4E, 0xC9, 0xB0)
MUTED = RGBColor(0x8A, 0x8E, 0xA8)

# Standard 16:9 in EMUs
WIDTH = Emu(12192000)
HEIGHT = Emu(6858000)

# --- Slide Content ---
SLIDES = [
    {
        "title": "The Baseload Problem",
        "bullets": [
            "Global electricity demand: 2x by 2050",
            "Renewables: intermittent, no 24/7 guarantee",
            "Conventional nuclear: $20\u201330B, 15+ years",
        ],
        "takeaway": "Can we get nuclear reliability without the cost?",
        "notes": (
            "By 2050, global electricity demand is projected to double. At the same time, "
            "we need to cut carbon emissions by roughly fifty percent this decade. Renewables "
            "are essential \u2014 but solar and wind are intermittent. They cannot guarantee power "
            "at 2am in January. Conventional nuclear can \u2014 but it costs twenty to thirty "
            "billion dollars per plant and takes over fifteen years to build. So the question "
            "is: can we get nuclear\u2019s reliability without its cost and complexity?"
        ),
    },
    {
        "title": "A New Manufacturing Model",
        "bullets": [
            "< 300 MW, factory-built modules",
            "Standardized components, shipped to site",
            "Incremental deployment as demand grows",
        ],
        "takeaway": "Not a smaller reactor \u2014 a different construction paradigm",
        "notes": (
            "That\u2019s the premise behind Small Modular Reactors. An SMR isn\u2019t just a smaller "
            "reactor \u2014 it\u2019s a fundamentally different manufacturing model. A conventional "
            "plant exceeds 1,000 megawatts and is built entirely on-site \u2014 custom-engineered, "
            "slow, and expensive. An SMR is under 300 megawatts, factory-fabricated under "
            "controlled conditions, then shipped and assembled on-site. That distinction is "
            "critical. Factory production means standardized components, tighter quality "
            "control, and learning-curve cost reductions \u2014 the same logic that transformed "
            "aerospace and shipbuilding. And crucially, you don\u2019t commit twenty billion "
            "dollars upfront. You deploy modules incrementally, matching capacity to demand."
        ),
    },
    {
        "title": "How It Works: Fission to Grid",
        "bullets": [
            "U-235 fission \u2192 E = mc\u00b2 energy release",
            "Heat \u2192 pressurized water \u2192 steam generator",
            "Rankine cycle: 30\u201335% thermal efficiency",
        ],
        "takeaway": "Not a design flaw \u2014 just physics",
        "notes": (
            "The underlying physics is identical to any nuclear plant. Uranium-235 undergoes "
            "fission \u2014 the nucleus splits, and because the fission products have slightly "
            "less mass than the original atom, that mass difference is converted to energy, "
            "exactly as E equals mc squared predicts. That energy heats pressurized water \u2014 "
            "held above 300 degrees Celsius by maintaining high pressure \u2014 which transfers "
            "thermal energy to a secondary loop through a steam generator. The steam drives "
            "a turbine connected to a generator. Thermodynamic efficiency sits at 30 to 35 "
            "percent \u2014 that\u2019s the Rankine cycle ceiling for this temperature range. Not a "
            "design flaw \u2014 just physics."
        ),
    },
    {
        "title": "NuScale: Passive Safety",
        "bullets": [
            "Natural circulation: no pumps needed",
            "Gravity-driven cooling \u2014 no power required",
            "Decay heat removed by thermodynamics alone",
        ],
        "takeaway": "The safety case is physics, not procedure",
        "notes": (
            "Now, here is where NuScale\u2019s design makes a genuinely important engineering "
            "contribution. Every nuclear reactor \u2014 even after shutdown \u2014 continues producing "
            "decay heat from fission product decay. In a conventional plant, you remove that "
            "heat with electrically powered pumps. But pumps need power, and power can fail. "
            "Fukushima showed exactly what happens when it does. NuScale eliminates that "
            "single point of failure entirely. Its cooling system relies on natural "
            "circulation: heated coolant is less dense, so it rises; cooler fluid descends "
            "to replace it. This is a convection loop driven by thermodynamics alone \u2014 no "
            "pumps, no external power, no operator action required. The safety case is "
            "physics, not procedure."
        ),
    },
    {
        "title": "Why NuScale Leads",
        "bullets": [
            "First & only NRC-certified SMR design",
            "Integral PWR: core + steam gen + pressurizer",
            "77 MWe per module \u2192 462 MW (6 modules)",
        ],
        "takeaway": "Not a concept \u2014 a certified, deployable design",
        "notes": (
            "NuScale is the most deployment-ready SMR design in the world. It is the first "
            "\u2014 and currently only \u2014 SMR to receive design certification from the U.S. "
            "Nuclear Regulatory Commission, arguably the most rigorous nuclear regulator on "
            "earth. Its integral pressurized water reactor design consolidates the reactor "
            "core, steam generator, and pressurizer into a single sealed pressure vessel \u2014 "
            "eliminating external piping and the failure points that come with it. Each "
            "module produces 77 megawatts electric. Six modules give you a 462 megawatt plant."
        ),
    },
    {
        "title": "The Economic Reality",
        "bullets": [
            "Capital cost: $6,000\u2013$10,000/kW installed",
            "LCOE: $65\u201390/MWh vs. solar at $30\u201340",
            "Idaho project cancelled \u2014 FOAK risk is real",
        ],
        "takeaway": "Loss of economies of scale is the core structural problem",
        "notes": (
            "Now the hard part \u2014 and I want to be honest about this. Capital costs for SMRs "
            "are currently estimated at six to ten thousand dollars per kilowatt installed. "
            "That is not cheaper than conventional nuclear \u2014 it\u2019s comparable or worse on a "
            "per-kilowatt basis. Projected levelized cost of energy sits at 65 to 90 dollars "
            "per megawatt-hour. That beats natural gas when you factor in carbon pricing, but "
            "it cannot compete with utility-scale solar at 30 to 40 dollars. The core problem "
            "is loss of economies of scale \u2014 larger plants spread fixed costs across more "
            "megawatts. SMRs don\u2019t get that advantage yet. And the cancellation of NuScale\u2019s "
            "flagship Idaho project confirmed that first-of-a-kind cost risk is real, not "
            "theoretical."
        ),
    },
    {
        "title": "Barriers to Deployment",
        "bullets": [
            "Regulatory: licensing not built for modular",
            "Supply chain: factory model unproven at scale",
            "Market: solar/wind costs still falling",
        ],
        "takeaway": "Nth-of-a-kind economics require the first plants to get built",
        "notes": (
            "Three structural barriers remain. First, regulatory: even with NRC "
            "certification, project-level licensing is slow and expensive \u2014 the regulatory "
            "framework was not designed for modular, repeated deployment. Second, supply "
            "chain: the factory manufacturing model that makes SMRs attractive doesn\u2019t exist "
            "at scale yet, so the cost advantages remain largely theoretical. Third, market "
            "competition: solar and wind costs keep falling. SMRs need to achieve "
            "nth-of-a-kind economics \u2014 where the tenth or twentieth unit benefits from "
            "accumulated manufacturing learning \u2014 to close that gap. That requires the first "
            "plants to actually get built."
        ),
    },
    {
        "title": "Where SMRs Fit",
        "bullets": [
            "Not a silver bullet \u2014 a specific tool",
            "Firm, dispatchable, low-carbon baseload",
            "Engineering proven; barrier is economic & solvable",
        ],
        "takeaway": "The grid needs what renewables alone cannot provide",
        "notes": (
            "So where does that leave us? Small Modular Reactors are not a silver bullet. "
            "They will not out-compete solar on cost, and they will not solve the storage "
            "problem that renewables face. But what they offer is something very specific: "
            "firm, dispatchable, low-carbon baseload power \u2014 the kind of generation the grid "
            "needs when the wind stops blowing and the sun goes down. The engineering is "
            "proven. The passive safety is real. The remaining barrier is economic \u2014 and it "
            "is solvable, if the first projects get financed and built."
        ),
    },
]


def add_text(slide, left, top, width, height, text, size, color,
             bold=False, align=PP_ALIGN.LEFT, name="Calibri"):
    """Add a simple text box."""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = name
    p.alignment = align
    return box


def set_bg(slide, color):
    """Set solid fill background."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def make_title_slide(prs):
    """Create the title slide using the built-in Title Slide layout."""
    layout = prs.slide_layouts[0]  # "Title Slide"
    slide = prs.slides.add_slide(layout)
    set_bg(slide, BG_DARK)

    # Use the built-in placeholders
    title_ph = slide.placeholders[0]
    subtitle_ph = slide.placeholders[1]

    # Title
    title_ph.text = "Small Modular Reactors"
    for p in title_ph.text_frame.paragraphs:
        p.font.size = Pt(44)
        p.font.color.rgb = WHITE
        p.font.bold = True
        p.font.name = "Calibri"
        p.alignment = PP_ALIGN.CENTER

    # Subtitle
    subtitle_ph.text = "Engineering, Economics, and the Case for NuScale"
    for p in subtitle_ph.text_frame.paragraphs:
        p.font.size = Pt(22)
        p.font.color.rgb = TEAL
        p.font.name = "Calibri Light"
        p.alignment = PP_ALIGN.CENTER

    # Hide date/footer/slide number placeholders
    for idx in (10, 11, 12):
        if idx in slide.placeholders:
            slide.placeholders[idx].text = ""

    # Notes
    notes = slide.notes_slide
    notes.notes_text_frame.text = "Title slide. Advance after brief introduction."

    return slide


def make_content_slide(prs, data, num, total):
    """Create a content slide using Title Only layout + text boxes."""
    layout = prs.slide_layouts[5]  # "Title Only"
    slide = prs.slides.add_slide(layout)
    set_bg(slide, BG_DARK)

    # Title via placeholder
    title_ph = slide.placeholders[0]
    title_ph.text = data["title"]
    for p in title_ph.text_frame.paragraphs:
        p.font.size = Pt(32)
        p.font.color.rgb = WHITE
        p.font.bold = True
        p.font.name = "Calibri"

    # Hide date/footer/slide number placeholders
    for idx in (10, 11, 12):
        if idx in slide.placeholders:
            slide.placeholders[idx].text = ""

    # Bullets as a single text box with multiple paragraphs
    bullet_box = slide.shapes.add_textbox(
        Inches(0.9), Inches(1.8), Inches(5.5), Inches(3.0)
    )
    tf = bullet_box.text_frame
    tf.word_wrap = True

    for i, bullet in enumerate(data["bullets"]):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = bullet
        p.font.size = Pt(20)
        p.font.color.rgb = LIGHT_GREY
        p.font.name = "Calibri Light"
        p.space_after = Pt(14)
        # Use bullet character
        p.text = "\u25b8  " + bullet

    # Takeaway line at bottom
    if data.get("takeaway"):
        add_text(slide, Inches(0.9), Inches(5.8), Inches(10), Inches(0.6),
                 data["takeaway"], 16, TEAL, name="Calibri Light")

    # Slide number
    add_text(slide, Inches(11.5), Inches(6.6), Inches(1.2), Inches(0.4),
             f"{num} / {total}", 11, MUTED, align=PP_ALIGN.RIGHT,
             name="Calibri Light")

    # Visual placeholder hint (right side) — simple text box
    add_text(slide, Inches(7.5), Inches(2.8), Inches(4.5), Inches(1.0),
             "[Add visual — see slides.md]", 14, MUTED,
             align=PP_ALIGN.CENTER, name="Calibri Light")

    # Speaker notes
    notes = slide.notes_slide
    notes.notes_text_frame.text = data["notes"]

    return slide


def main():
    prs = Presentation()
    prs.slide_width = WIDTH
    prs.slide_height = HEIGHT

    make_title_slide(prs)

    total = len(SLIDES)
    for i, data in enumerate(SLIDES, 1):
        make_content_slide(prs, data, i, total)

    out = "SMR_NuScale_Presentation.pptx"
    prs.save(out)
    print(f"Saved: {out} ({total + 1} slides)")


if __name__ == "__main__":
    main()
