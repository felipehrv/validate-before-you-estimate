"""
Build the Word documents for the mock lesson:
  handouts/Handout_Validate_Before_You_Estimate.docx   (participants, 2 pages)
  docs/Lesson_Plan_Validate_Before_You_Estimate.docx   (one page, for the panel)
  docs/Talking_Points_and_QA.docx                      (for Felipe, bullets only)
Run: python3 scripts/06_build_docs.py
"""

import json
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

PROJECT = Path(__file__).resolve().parents[1]
R = json.load(open(PROJECT / "outputs/results.json"))
X = json.load(open(PROJECT / "outputs/exercise_items.json"))
(PROJECT / "handouts").mkdir(exist_ok=True)
(PROJECT / "docs").mkdir(exist_ok=True)

INK = RGBColor(0x1B, 0x1F, 0x23)
MUTED = RGBColor(0x52, 0x51, 0x4E)
BLUE = RGBColor(0x2A, 0x78, 0xD6)
ORANGE = RGBColor(0xEB, 0x68, 0x34)

pct = lambda v, d=1: f"{100 * v:.{d}f}%"
pts = lambda v, d=1: f"{100 * v:.{d}f}"
E = R["estimates"]
S = {f"{r['model']}|{r['prompt']}": r for r in R["showdown"]}
BP = R["by_party_haiku_p3"]
BE = R["by_era_haiku_p3"]
LP = R["linkage_policies"]
MV = R["moved_rate"]
PJ = R["sample_size_curve_projected"]
PJrow = {r["n_hand_coded"]: r for r in PJ["rows"]}
N_C3 = -(-PJ["n_for_target_corrected"] // 50) * 50        # rounded up to the next 50
N_H3 = -(-PJ["n_for_target_human_only"] // 50) * 50
h1, h2, h3, s3 = (S["claude-haiku-5-5|prompt_1_names"], S["claude-haiku-5-5|prompt_2_codebook"],
                  S["claude-haiku-5-5|prompt_3_codebook_examples"], S["claude-sonnet-5-5|prompt_3_codebook_examples"])
GAP = BP["Democrat"]["health_llm"] - BP["Republican"]["health_llm"]


def base_doc(margins=0.8, font="Calibri", size=11):
    doc = Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = Inches(margins)
        s.left_margin = s.right_margin = Inches(margins)
    st = doc.styles["Normal"]
    st.font.name = font
    st.font.size = Pt(size)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), font)
    st.paragraph_format.space_after = Pt(4)
    return doc


def heading(doc, text, size=16, color=INK, before=10, after=4, bold=True):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    r.font.color.rgb = color
    return p


def para(doc, text, size=11, color=INK, italic=False, bold=False, after=4, align=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.italic = italic
    r.bold = bold
    return p


def bullet(doc, text, size=11, level=0, bold_lead=None, after=3):
    p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
    p.paragraph_format.space_after = Pt(after)
    if bold_lead:
        r = p.add_run(bold_lead)
        r.bold = True
        r.font.size = Pt(size)
    r = p.add_run(text)
    r.font.size = Pt(size)
    return p


def shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def table(doc, rows, widths, header=True, size=10, fill="2F3A44"):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    tblPr = t._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    for j, wdt in enumerate(widths):
        t.columns[j].width = Inches(wdt)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = t.cell(i, j)
            c.width = Inches(widths[j])
            c.text = ""
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(str(val))
            r.font.size = Pt(size)
            if header and i == 0:
                r.bold = True
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                shade(c, fill)
    return t


# ===========================================================================
# 1. HANDOUT
# ===========================================================================
doc = base_doc(margins=0.7, size=10.5)
heading(doc, "Validate Before You Estimate", size=20, before=0, after=0)
para(doc, "Handout: three challenges and one question", size=12, color=MUTED, after=8)

heading(doc, "Challenge 1: be the coder (3 minutes)", size=13, color=BLUE, before=4)
para(doc, "Code each bill title into exactly one topic using the codebook below. Choose the topic that best describes the main purpose of the bill; when a bill touches several topics, code its primary target. Then compare with your neighbor and circle the titles where you disagree.", after=6)

cb = [
    ("Health", "health care reform; insurance, Medicare and Medicaid; drug industry and approval; hospitals and clinics; health workforce; disease prevention and public health; mental health; long-term care; tobacco; treatment and prevention of drug and alcohol abuse; medical research"),
    ("Education", "higher education and student aid; elementary and secondary schools; underprivileged, vocational and special education; standards; education research"),
    ("Environment", "drinking water; waste and hazardous waste; air pollution and climate; recycling; species and forests; land and water conservation. Not energy, farming or the management of parks and public lands"),
    ("Immigration", "immigration, naturalization and citizenship; refugees and asylum; visas; enforcement of immigration law"),
    ("Law and Crime", "law enforcement agencies; white collar crime; drug trafficking and enforcement; courts; prisons and sentencing; juvenile crime; child abuse; family law and domestic violence; criminal and civil code; police"),
    ("Social Welfare", "low-income assistance including food assistance; the elderly including Social Security retirement; people with disabilities; charities and volunteers; child care"),
    ("Other", "everything else: economy and taxes, civil rights, agriculture, labor, energy, transportation, housing, commerce, defense and veterans, science, trade, foreign affairs, government operations, commemorations, public lands, culture, relief of named individuals"),
]
rows = [["Topic", "Codebook (condensed from the Comparative Agendas Project)"]] + [[a, b] for a, b in cb]
table(doc, rows, [1.2, 5.9], size=9)
doc.add_paragraph().paragraph_format.space_after = Pt(2)

rows = [["#", "Bill title", "Your code", "Neighbor's code"]]
for i, h in enumerate(X["hand_coding"], 1):
    rows.append([str(i), f"({h['year']}) {h['title']}", "", ""])
table(doc, rows, [0.3, 5.0, 0.9, 0.9], size=9)
para(doc, "Bell question: on which title did you and your neighbor disagree, and what rule would have settled it?", size=10, italic=True, color=MUTED, after=2)

doc.add_page_break()
heading(doc, "Challenge 2, part one: place your bet", size=13, color=BLUE, before=0)
para(doc, "The same 1,650 titles were coded three times by the same cheap model with three prompts. Which one agrees most with the human coders? Tick one, then write a number.", after=4)
rows = [["", "Instrument", "Your guess of its accuracy"],
        ["[  ]", "P1  Category names only (no definitions)", "____ %"],
        ["[  ]", "P2  The condensed codebook", "____ %"],
        ["[  ]", "P3  The codebook plus 14 human-coded examples", "____ %"]]
table(doc, rows, [0.5, 4.3, 2.3], size=10)
para(doc, "Second bet: a stronger model with P3 beats the best cheap prompt by  [  ] less than 5 points   [  ] more than 5 points.   Third bet: a lazy coder who answers \"Other\" every time scores ____ % accuracy.", size=10, after=8)

heading(doc, "Challenge 2, part two: whose error is it?", size=13, color=BLUE, before=6)
para(doc, "For each title the model and the human coder disagree. Vote, then say what you would change (a prompt rule, a codebook clarification, or the human code).", after=4)
rows = [["", "Title", "Human", "Model", "Who is wrong?"]]
for i, d in enumerate(X["disputed"][:3]):
    rows.append([chr(65 + i), f"({d['year']}) {d['title']}", d["gold"], d["model"], "[ ] model  [ ] human  [ ] codebook"])
table(doc, rows, [0.3, 3.9, 0.9, 0.9, 1.1], size=8.5)
doc.add_paragraph().paragraph_format.space_after = Pt(2)

heading(doc, "Challenge 3: how many titles would you code by hand? (exit ticket)", size=13, color=BLUE, before=6)
para(doc, f"Your project: the same instrument on 20,000 bill titles. Your committee wants to know whether Democrats and Republicans differ by 3 points or more in the share of health bills. A research assistant codes about 100 titles an hour. Rule: to tell a 3-point gap between two groups from zero, each group's interval must be about plus or minus 2 points or narrower. The table is the curve on the screen (this lesson's error rates, projected to a group of {PJ['n_group']:,} titles). It applies to each group separately.", size=10, after=4)
sizes = [50, 100, 200, 300, 400, 800]
rows = [["Hand-coded titles in the group (n)"] + [str(n) for n in sizes],
        ["Corrected interval, plus or minus"] + [pct(PJrow[n]["ci_halfwidth"]) for n in sizes],
        ["Hand-coded sample alone, plus or minus"] + [pct(PJrow[n]["human_only_halfwidth"]) for n in sizes]]
table(doc, rows, [2.4, 0.78, 0.78, 0.78, 0.78, 0.78, 0.78], size=9)
para(doc, "I would hand-code ______ titles per party, ______ titles in all, which is ______ RA hours. I would promise my committee an interval of about plus or minus ______ points per party. Without the model I would need ______ per party.", size=10, after=8)

heading(doc, "Question (part two): would you rather lose pairs or invent them?", size=13, color=BLUE, before=6)
para(doc, "Linking two files of people with no shared ID. The research question is the share of people whose postcode changed. Mark your choice and the reason.", after=4)
rows = [["Policy for the ambiguous pairs", "Precision", "Recall", "Estimated share who moved", "Your choice"],
        ["Drop them (strict)", pct(LP["Rules, strict (drop the band)"]["precision"]), pct(LP["Rules, strict (drop the band)"]["recall"]), pct(MV["Rules, strict"]), "[  ]"],
        ["Keep them all (lenient)", pct(LP["Rules, lenient (keep the band)"]["precision"]), pct(LP["Rules, lenient (keep the band)"]["recall"]), pct(MV["Rules, lenient"]), "[  ]"],
        ["Let a model adjudicate", pct(LP["Rules + Haiku on the band"]["precision"]), pct(LP["Rules + Haiku on the band"]["recall"]), pct(MV["Rules + Haiku"]), "[  ]"],
        ["Truth (every true link known)", "", "", pct(MV["Truth (all true links)"]), ""]]
table(doc, rows, [2.3, 1.0, 1.0, 1.9, 0.9], size=9)
para(doc, "Precision: when the rule says same person, how often it is right. Recall: of the true links, how many it finds. Materials, data and the notebook: Felipe H. Rodriguez Villalta, felipe.hrv@gmail.com", size=9, color=MUTED, after=0)
doc.save(PROJECT / "handouts/Handout_Validate_Before_You_Estimate.docx")
print("handout written")

# ===========================================================================
# 2. LESSON PLAN (one page)
# ===========================================================================
doc = base_doc(margins=0.6, size=9.5)
heading(doc, "Lesson plan: Validate Before You Estimate", size=16, before=0, after=0)
para(doc, "Using AI to label text and link records in social science research. 45 minutes, graduate students in the social sciences and humanities. Felipe H. Rodriguez Villalta, prepared for the D-Lab, UC Berkeley, November 2026.", size=9.5, color=MUTED, after=6)

heading(doc, "Learning objectives", size=11, color=BLUE, before=2, after=2)
for t in [
    "Explain why a classification prompt is a measurement instrument, and what changes when the codebook goes into it.",
    "Design a validation sample and read accuracy, kappa, precision and recall as measurement error.",
    "Explain why high accuracy does not guarantee an unbiased estimate or a valid comparison between groups.",
    "Read a bias-corrected estimate (prediction-powered inference) that combines AI labels with a small hand-coded sample, and choose how many documents to hand-code from an interval curve.",
    "Apply the same recipe to record linkage: rules for the easy pairs, an AI adjudicator for the ambiguous band, a reviewed sample to measure the error.",
]:
    bullet(doc, t, size=9.5, after=1)

heading(doc, "Prerequisites, materials, fit", size=11, color=BLUE, before=6, after=2)
bullet(doc, "None for the room exercises (hand-coding, voting, reading a curve); Python Fundamentals recommended to run the notebook afterward; no AI experience required.", size=9.5, after=1)
bullet(doc, "Data: Comparative Agendas Project congressional bills (1,650 titles, 75 per Congress 1973 to 2016, human-coded); Febrl person records (synthetic, 5,000 true links known). All model outputs pre-computed and cached (6,600 labels, 2,966 pair decisions; model, prompt version, tokens and timestamp logged). Runs on DataHub, Binder or a laptop without an API key.", size=9.5, after=1)
bullet(doc, "Catalog fit: builds on Python APIs for Large Language Models (mechanics) and Prompt Engineering; sits next to Qualitative Coding with AI Tools (the passage-by-passage audit) as the measurement step; leads into a full Record Linkage workshop. Not covered: model internals, fine-tuning, inductive coding, agent pipelines. The repository follows the D-Lab template so a Data Science Fellow can teach it after one rehearsal.", size=9.5, after=1)

heading(doc, "Sequence", size=11, color=BLUE, before=6, after=2)
rows = [["Min", "Block", "Instructor", "Participants (exercise goal)"],
        ["0-5", "Hook and poll", "A trend chart built from AI labels; Desai, Card and Jacobs (2026) on validation practice; the one idea", "Poll by hand: used an LLM to code? checked it? did the check change a number?"],
        ["5-13", "Codebook to prompt", "Three instruments for one task; structured output; logging; nothing in a prompt is scored", "Challenge 1: code six titles by hand, compare with a neighbor (feel the ambiguity; set the human ceiling)"],
        ["13-24", "Validate", f"Accuracy and kappa for P1, P2, P3 ({pct(h1['accuracy'],0)}, {pct(h2['accuracy'],0)}, {pct(h3['accuracy'],0)}); confusion matrix; precision and recall by class; error by party and era", "Challenge 2: bet on the winning prompt; then vote on three disputed titles (model wrong, human wrong, or codebook ambiguous) and say what to change"],
        ["24-32", "Estimate", f"Naive vs hand-coded vs corrected share of health bills; a {pts(GAP)}-point partisan gap the humans do not see dissolves after correction; the scale analogy; the PPI line; the era chart and what 63 to 72 titles per era can and cannot say", f"Challenge 3 (exit ticket): read the interval curve and choose how many titles to hand-code per group to see a 3-point gap (checkable answer: about {N_C3} per group, {N_H3} without the model)"],
        ["32-40", "Link records", f"Same recipe: rules sort pairs, the model adjudicates the band; precision and recall of four policies; the downstream mobility rate ({pct(MV['Rules, strict'])} to {pct(MV['Rules, lenient'])} vs truth {pct(MV['Truth (all true links)'])})", "Question: lose pairs or invent them? Decide for the number you want to know"],
        ["40-45", "Close", "What we know works; what to report; privacy and cost; key points", "Take-home challenge on their own data; exit tickets collected"]]
table(doc, rows, [0.5, 1.0, 3.0, 2.8], size=8.5)

heading(doc, "Teaching approach and assessment", size=11, color=BLUE, before=6, after=2)
bullet(doc, "Backward design from the decisions a student makes in a real project; predict-observe-explain before every result is shown; an exercise every ten to twelve minutes (the D-Lab guide asks for at least one every fifteen); colloquial language, each term explained when first used; it is OK not to know, said out loud.", size=9.5, after=1)
bullet(doc, "Formative assessment: opening poll, bell questions as attention checks, Challenge 3 as an exit ticket with a checkable number; the take-home challenge links to D-Lab consulting.", size=9.5, after=1)
bullet(doc, "Ethics and evaluation inside the method: differential error by group, what data may leave the machine, cost and time, reporting standards.", size=9.5, after=1)
bullet(doc, "Contingencies: cached outputs and a saved dry run replace any live call; the handout makes the exercises work with or without laptops. The 40-minute mock version (overview plus lesson) folds the precision-recall and reporting slides into their neighbors and takes a minute off each challenge; a 35-minute lesson also drops the four-policies chart. The era chart always stays, because it answers the opening hook.", size=9.5, after=1)
doc.save(PROJECT / "docs/Lesson_Plan_Validate_Before_You_Estimate.docx")
print("lesson plan written")

# ===========================================================================
# 3. TALKING POINTS + Q&A (for Felipe)
# ===========================================================================
doc = base_doc(margins=0.8, size=11)
heading(doc, "Talking points and Q&A: Validate Before You Estimate", size=18, before=0, after=0)
para(doc, "D-Lab mock lesson, Thursday November 5, 2026, 2:40 PM. Target 40 minutes, then questions. Bullets only, by slide, with the clock time at which each slide should start and the number to say.", size=11, color=MUTED, after=8)

heading(doc, "The clock, and what to skip if you are behind", size=13, color=BLUE, before=6, after=3)
bullet(doc, "D-Lab's window is 35 to 45 minutes for overview plus lesson. The plan below is 40: overview 0:00 to 8:00 (slides 1 to 8), part one 8:00 to 32:00 (slides 9 to 23, three challenges), part two 32:00 to 37:00 (slides 24 to 26), close 37:00 to 40:00 (slides 27 to 30). Five minutes of slack before 45; never pass it.", after=2)
bullet(doc, "Two slides are folded by default in the mock: slide 17 (precision and recall) is one sentence said over the heatmap, and slide 28 (report, privacy, cost) is thirty seconds. Both exist in full for the 45-minute lesson.", after=2)
bullet(doc, "Checkpoint at the end of Challenge 2b (slide 18): you should be at 23:00. If you are past 24:30, say slide 20 (the formula) in one sentence and read only the two bold bullets on slide 21.", after=2)
bullet(doc, "Checkpoint at the end of Challenge 3 (slide 23): you should be at 32:00. If you are past 33:30, skip slide 25 (say the three policies' numbers on slide 26) and slide 28.", after=2)
bullet(doc, "Never skip a challenge. If everything else fails, Challenge 1 becomes two minutes and Challenge 3 three.", after=2)

def section(title):
    heading(doc, title, size=13, color=BLUE, before=10, after=3)

def slide(n, title, minute, items):
    heading(doc, f"Slide {n}  |  {title}  |  {minute}", size=11.5, before=6, after=2)
    for it in items:
        if isinstance(it, tuple):
            bullet(doc, it[1], bold_lead=it[0], after=2)
        else:
            bullet(doc, it, after=2)

section("Part A. Overview (slides 1 to 8, 0:00 to 8:00)")
slide(1, "Title", "0:00", [
    "One line on who you are: research scientist, ten years making other people's research data usable (Stanford, ILO, World Bank, UCSF), and you teach the way you work.",
    "What happens today: the overview of the full 45-minute lesson first (objectives, exercises, approach, where it sits), then you teach the core with the panel as students.",
    "Hand out the sheet now so it is on the table when Challenge 1 arrives.",
])
slide(2, "Why this lesson, and why now", "0:45", [
    ("The practice: ", "students already code interviews, open answers, documents and posts with LLMs; cheap, fast, often good."),
    ("The gap: ", "Desai, Card and Jacobs (2026) reviewed 27 papers in eight flagship social science journals: LLM measurements central to the analysis, validation limited and inconsistent."),
    ("The risk: ", "Egami and coauthors: treating predicted labels as true biases estimates and breaks confidence intervals even above 90% accuracy."),
    ("The one idea: ", "every AI label is a measurement with error; measure the error against people, then carry it into the estimate. Everything else is this sentence applied twice."),
])
slide(3, "Learning objectives", "2:00", [
    "Read the five quickly; stress the verbs: explain, design, explain, read, apply. In the notebook, read becomes run.",
    "Objective 3 is the heart: high accuracy does not guarantee an unbiased estimate. Say that most AI workshops stop at objective 2.",
])
slide(4, "The lesson at a glance", "3:00", [
    "Timeline: hook 5, codebook to prompt 8, validate 11, estimate 8, link records 8, close 5.",
    "Four exercises, one every ten to twelve minutes (the D-Lab guide asks for at least one every fifteen); each goal is phrased as what the participant can do afterward.",
    "Challenge 1 sets the human ceiling; Challenge 2 separates prompt from model, then decides whose error it is; Challenge 3 is the budget decision with a checkable answer; the question in part two makes linkage an estimation choice.",
])
slide(5, "Teaching approach", "4:15", [
    ("Backward design: ", "started from the decisions a student makes in a real project (how many to hand-code, what to report) and built exercises for those decisions."),
    ("Predict, observe, explain: ", "a commitment before any result is shown; the gap between the guess and the data is what sticks."),
    ("Hands-on without a key: ", "every model call pre-run and cached; the notebook runs on DataHub or Binder with no API key, or on a laptop with no internet; a dry-run file shows what one call sends and returns."),
    ("A path without code: ", "hand-coding, voting and reading a curve need no Python; a student from history or anthropology completes the lesson."),
    ("Ethics and evaluation inside the method: ", "differential error, privacy of data sent to an API, cost, reporting; steps of the workflow, not a closing slide."),
    ("Formative assessment: ", "poll, bell questions, Challenge 3 as an exit ticket with a checkable number."),
])
slide(6, "Where it sits in the D-Lab catalog", "5:45", [
    "Python APIs for LLMs teaches the mechanics (structured output via OpenRouter). Qualitative Coding with AI Tools teaches the audit: code six passages by hand, run the agent twice, check the codes against the text. This lesson is the measurement step between them: the error at scale, carried into the estimate.",
    "What is new: the method is from Berkeley (PPI) and Princeton (DSL); the 45-minute no-code version, the hand-coding budget decision and the linkage half are the lesson's own.",
    "The next workshop is Record Linkage for Research in full: blocking, comparison, probabilistic matching, clerical review, privacy; AI as one adjudicator.",
    "Builds on Python APIs for LLMs and Prompt Engineering; Python Fundamentals recommended to run the notebook afterward.",
    "Fits a 50-minute AI Pulse session or part 3 of a two-hour workshop.",
])
slide(7, "Materials", "6:45", [
    "CAP bills: 1,650 titles, 75 per Congress 1973 to 2016, human-coded (90 to 95% coder agreement). Febrl: 5,000 person records twice, synthetic, true links known.",
    "6,600 labels and 2,966 pair decisions pre-computed; every label logged with model, prompt version, tokens, timestamp. Whole run under 20 dollars.",
    "Repo on the D-Lab template: README, lessons and solutions notebooks, data, prompts, dry run. Runs on DataHub and Binder, so a Data Science Fellow can teach it after one rehearsal. Handout for the room. Thirty seconds, no more; the handout is already on the table.",
])
slide(8, "Transition", "7:30", [
    "Change register: slower, eye contact. \"For the next thirty minutes you are my graduate students.\"",
    "Three questions they will answer: can I trust this label, can I trust this estimate, how many do I code by hand.",
    "Say the IOKN2K line out loud: it is OK not to know; every bet is a guess, and the wrong guesses teach the most.",
])

section("Part B. Lesson, part one: labels as measurements (slides 9 to 23, 8:00 to 32:00)")
slide(9, "Hook", "8:00", [
    f"Show the chart: health bills from {pct(BE['1973-1986']['health_llm'])} to {pct(BE['2001-2016']['health_llm'])} of Congress's agenda, climbing era after era. Ask: would you put this in your dissertation?",
    "Poll by hand: used an LLM to code text? checked it against a person? did the check change a number? Hands drop at question three; that gap is the lesson.",
    "The chart is real, labels from a model with 88% accuracy. By the end they will know how much of the climb is the bills and how much is the instrument. Do not answer it yet.",
])
slide(10, "The data", "9:30", [
    "CAP: every U.S. bill since 1947 coded into 21 topics by trained coders. We keep seven labels.",
    "1,650 titles, 75 per Congress, with the sponsor's party and chamber.",
    "\"Other\" is 72% of bills, so accuracy alone will flatter any classifier. Human coders agree 90 to 95%: even gold has error.",
    "Say the 'your version' line looking at the room: transcripts coded for housing insecurity by gender, newspaper ads coded for occupation by decade, clinic notes coded for tobacco use by site. It tells a sociologist, a historian and a public health student that this is their problem.",
])
slide(11, "Codebook to prompt", "10:30", [
    "Three instruments: P1 names only; P2 the condensed codebook with boundary rules (drug treatment is Health, trafficking is Law and Crime, veterans are Other); P3 the codebook plus 14 human-coded examples from outside the evaluation sample.",
    "Fixed JSON answer with a one-sentence reason; logged with model, prompt version, tokens, timestamp. The prompt is frozen before scoring.",
    "The glossary line at the bottom is there for the anthropologist in the room: JSON, tokens, Haiku and Sonnet, reasoning off. Read it only if faces ask.",
    "The word instrument is deliberate: a changed prompt is a changed instrument and gets its own validation.",
])
slide(12, "Challenge 1: be the coder", "12:00", [
    "Three real minutes. Walk around. Then ask two pairs where they disagreed before showing the results.",
    "Items: controlled substances take-back (Health), law enforcement retirement (Law and Crime), FICA exemption over 65 (Social Welfare), zone academy bonds (Education), border security (Immigration), national heritage areas (Other).",
])
slide(13, "What the coders said", "15:00", [
    "Where they disagreed with their neighbor is where the model will disagree with the humans. Human agreement is the ceiling, not 100%.",
    "Titles 3 and 4 change the tax law; the codebook lists taxes under Other, but the CAP coders put them under Social Welfare and Education, following the rule 'code the bill by who it is for, not by the tool it uses'. Most people miss that rule, and so did the model. That is a codebook fix: make the rule explicit in the prompt and re-validate.",
    "The answer key is a measurement too: a real case from the CAP data, a bill funding treatment centers for torture victims, is coded Environment. Keep it for Challenge 2: sometimes the human slipped.",
])
slide(14, "Challenge 2, part one: place your bet", "16:30", [
    "Same titles, same cheap model, three prompts. Show of hands per prompt. Second bet: does Sonnet with P3 beat Haiku with P3 by more or less than five points? Third bet: accuracy of a lazy coder who always answers Other.",
])
slide(15, "The results", "17:30", [
    ("Numbers: ", f"P1 {pct(h1['accuracy'])}, P2 {pct(h2['accuracy'])}, P3 {pct(h3['accuracy'])}; Sonnet P3 {pct(s3['accuracy'])}. Kappa {h1['kappa']:.2f}, {h2['kappa']:.2f}, {h3['kappa']:.2f}, {s3['kappa']:.2f}."),
    f"The codebook in the prompt is worth {pts(h2['accuracy']-h1['accuracy'],0)} points; the examples add {pts(h3['accuracy']-h2['accuracy'])}, which is inside run-to-run noise (a paired test on the same 1,650 titles cannot tell P2 from P3); the model upgrade {pts(s3['accuracy']-h3['accuracy'])}, and that one is real (p about 0.01). Say points, not percent.",
    "Always-Other scores 72% and kappa 0.00. Kappa asks how much better than chance agreement the instrument is (chance from both coders' label frequencies, not the 72%), 0 to 1; report it next to accuracy.",
    "The instrument is what you control and what you must report.",
])
slide(16, "Where the disagreements live", "19:00", [
    f"Read the matrix by row (rows are the human codes). Health row: {R['confusion_haiku_p3']['Health']['Health']} health bills the model also called Health, {R['confusion_haiku_p3']['Other']['Health']} it called Other. Other row: {R['confusion_haiku_p3']['Health']['Other']} bills the humans coded Other that the model called Health.",
    f"Social Welfare keeps {pct(R['per_class_haiku_p3'][5]['recall'],0)} of its bills, Education {pct(R['per_class_haiku_p3'][1]['recall'],0)}: small categories pay for the big one's accuracy.",
    "Most disagreements involve Other, in both directions: Other is the catch-all category, and the line between Other and a specific topic is the hardest to draw.",
])
slide(17, "Precision and recall (folded: one sentence, no stop)", "20:00", [
    f"Two questions per category. Precision: when the model says Health, how often the humans agree ({pct(R['per_class_haiku_p3'][0]['precision'],0)}). Recall: when the humans say Health, how often the model notices ({pct(R['per_class_haiku_p3'][0]['recall'],0)}).",
    "Say it in one breath and move on; the handout repeats the two words in the linkage question. In the 45-minute lesson this slide gets a minute.",
])
slide(18, "Challenge 2, part two: whose error", "20:30", [
    ("A, Defense Health Program appropriation: ", "humans Health, model Other (it followed the word defense). A prompt rule would fix it: health programs inside any department are Health."),
    ("B, cigarette fire-safety standards: ", "humans Other (consumer product safety), model Health (it heard tobacco). Codebook clarification: tobacco means use and health effects, not product safety."),
    ("C, Social Security marriage rule for divorced spouses: ", "humans Law and Crime (family law), model Social Welfare. Reasonable either way; the codebook's own structure is counterintuitive here."),
    "After the vote, give the results: A, the human is right and the model wrong (fix: a prompt rule). B, the model is wrong (fix: a codebook clarification). C, the codebook is ambiguous (fix: decide and write it down). Three owners of error, three fixes. Let them argue for two minutes. Checkpoint: 23:00 when you leave this slide.",
])
slide(19, "The error is not the same for everyone", "23:00", [
    ("Counted with the human codes: ", f"{pct(BP['Democrat']['health_gold'])} of the bills sponsored by Democrats and {pct(BP['Republican']['health_gold'])} of those sponsored by Republicans are about health. No gap."),
    ("Counted with the model's labels, same bills: ", f"{pct(BP['Democrat']['health_llm'])} vs {pct(BP['Republican']['health_llm'])}. A {pts(GAP)}-point gap that does not exist. The model created it."),
    f"Accuracy is almost the same for both parties ({pct(BP['Democrat']['accuracy'],0)} vs {pct(BP['Republican']['accuracy'],0)}); the sign of the error differs, and that bends a comparison.",
    f"By era the error changes sign: over-count in 1973-86 and 2001-16, under-count in between; accuracy slips from {pct(BE['1973-1986']['accuracy'],0)} to {pct(BE['2001-2016']['accuracy'],0)}. The humans' series levels off, the model's keeps climbing. Always validate by the groups you will compare.",
])
slide(20, "Measure the error, then subtract it", "24:30", [
    "The scale analogy: a scale reads heavy. Weigh a few people on a good scale too, measure how far off it is on average, subtract that from the group's average. You correct the average, not each person.",
    "Corrected share = mean of model labels over all N titles, minus the mean of (model minus human) over the n hand-coded titles. A label is a 1 or a 0; the interval says how much n titles can tell you.",
    "Prediction-powered inference (Angelopoulos, Bates, Fannjiang, Jordan, Zrnic, Science 2023, Berkeley) and design-based supervised learning (Egami, Hinck, Stewart, Wei). One line; it holds whatever the error pattern, as long as the subset is a random sample.",
])
slide(21, "Three estimates of one share", "26:00", [
    "Ask first: which way will the correction move the Democratic estimate?",
    "Optional, your call: switch to the solutions notebook on your laptop and re-run the two cells that print the three estimates by party (cached data, no API call, deterministic; about 60 seconds). It shows the one line of arithmetic running, which is how the D-Lab teaches. If the screen or the adapter fails, the slide has the same numbers; do not spend a minute fighting it.",
    ("Naive: ", f"{pct(E['Democrat']['naive']['est'])} vs {pct(E['Republican']['naive']['est'])}, a {pts(E['Democrat']['naive']['est'] - E['Republican']['naive']['est'])}-point gap with intervals that ignore the model's error. On 1,650 titles the two intervals overlap, so call it marginal; on 20,000 documents the same gap would sit about six standard errors from zero, far outside any naive interval, because naive intervals shrink with N and the bias does not. That is the sentence to land."),
    ("Hand-coded subset only (200): ", f"{pct(E['Democrat']['human']['est'])} vs {pct(E['Republican']['human']['est'])}. Right on average; any one sample of 200 can land a point or two off, which is what the wide interval says."),
    ("Corrected: ", f"{pct(E['Democrat']['corrected']['est'])} vs {pct(E['Republican']['corrected']['est'])}, with an interval about {E['Democrat']['corrected']['se']/E['Democrat']['human']['se']*100:.0f}% as wide as the subset-only one. The gap is gone."),
    "The black mark is all 1,650 hand-coded titles: the truth we usually cannot afford, here it checks the method.",
])
slide(22, "Back to the opening chart", "27:30", [
    f"Humans: {pct(E['1973-1986']['truth'])}, {pct(E['1987-2000']['truth'])}, {pct(E['2001-2016']['truth'])}. Model: {pct(E['1973-1986']['naive']['est'])}, {pct(E['1987-2000']['naive']['est'])}, {pct(E['2001-2016']['naive']['est'])}. Do not claim the humans' dip after 2000 or the model's climb on their own (each is within noise for 525 to 600 titles); claim the divergence: on the same bills the model's net error flips from minus 1 point to plus 2 after 2000, which is well beyond noise. The climb after 2000 is the instrument.",
    f"Be honest before anyone asks: with {E['1973-1986']['n_hand']} to {E['1987-2000']['n_hand']} hand-coded titles per era the corrected intervals are plus or minus 4 to 5 points; they cover the truth everywhere but cannot tell a plateau from a climb, and in 1973-86 the corrected point lands farther from the truth than the naive one. The subset alone found 1 health bill in 63: its interval runs from {pct(E['1973-1986']['human']['ci_low'])} to {pct(E['1973-1986']['human']['ci_high'])} (a Wilson interval; the textbook formula dips below zero near 0 and is too narrow). The interval that admits its width is the answer; the point alone would mislead.",
    "That is the price of splitting 200 titles three ways, and the bridge to Challenge 3.",
])
slide(23, "Challenge 3: how many to hand-code (exit ticket)", "28:30", [
    "Read the scenario from the slide: 20,000 titles, a 3-point question between the parties, 100 titles an hour, and the rule that each group needs plus or minus 2 or narrower. Three minutes to write n per party, hours, and the interval on the handout.",
    (f"The checkable answer: ", f"about {N_C3} per party ({pct(PJrow[400]['ci_halfwidth'])} at 400), {2 * N_C3} titles, {2 * N_C3 // 100} RA hours, one day. Hand-coding alone: about {N_H3} per party, {2 * N_H3 // 100} hours. With 200 per party the interval is plus or minus {pct(PJrow[200]['ci_halfwidth'])}: a 3-point gap could not be told from zero."),
    "The lesson's floor (200 at random) is the wrong answer to this question, and that is the point: the floor is for one share, a comparison costs more per group. The curve comes from this lesson's error rates; in their project it comes from a pilot of 100 hand-coded documents.",
    "Doubling n cuts the width by about 30%; four times the work buys half. Fewer than 200 documents in the whole corpus? Code them all. Checkpoint: 32:00 when you leave this slide.",
    "If a methods person pushes: plus or minus 2 per group makes a 3-point gap just distinguishable from zero (about 55% power). To be confident of catching it (80% power) each group needs about plus or minus 1.5: about 720 per party with the model, about 1,500 by hand. Same curve, stricter target; the exercise's logic is unchanged.",
])

section("Part C. Lesson, part two: links as measurements (slides 24 to 26, 32:00 to 37:00)")
slide(24, "Second problem, same recipe", "32:00", [
    "Febrl: file B is a corrupted copy of file A (typos, swapped digits, moved addresses); synthetic, so the 5,000 true links are known. Four shared fields, no ID. Your version: census manuscripts across decades, hospital discharges to death records.",
    "One clause for the methods people: at scale the rules become Fellegi-Sunter probabilistic weights, the band is their uncertain middle, and deduplication within one file is the same problem. The recipe does not change.",
    f"Comparing everyone with everyone would be 25 million pairs; blocking compares only pairs that share something: {R['blocking']['candidates']:,} candidate pairs, {R['blocking']['lost_links']} true links lost, and all {R['blocking']['lost_links']} are movers (the lesson's mechanism is already at work in the blocking step; say it, it is a nice touch). Rules sort {R['blocking']['sure']:,} sure matches and leave a band of {R['blocking']['band']:,} ambiguous pairs with {R['blocking']['band_true']} true links inside.",
    "Linking is classification too: the unit is the pair, the label is same person or not.",
])
slide(25, "Four ways to handle the band", "33:30", [
    ("Drop the band: ", f"precision {pct(LP['Rules, strict (drop the band)']['precision'])}, recall {pct(LP['Rules, strict (drop the band)']['recall'])}; 14% of the people are gone."),
    ("Keep the band: ", f"recall {pct(LP['Rules, lenient (keep the band)']['recall'])}, precision {pct(LP['Rules, lenient (keep the band)']['precision'])}; one in six links is invented."),
    ("Model adjudicates: ", f"precision {pct(LP['Rules + Haiku on the band']['precision'])}, recall {pct(LP['Rules + Haiku on the band']['recall'])}; inside the band it finds {pct(R['band_claude-haiku-5-5']['recall_in_band'],0)} of true pairs and is right {pct(R['band_claude-haiku-5-5']['precision_in_band'],0)} of the time it says yes. The stronger model is not better here."),
    "Still a measurement: in a real project you hand-review a random sample of the band (precision) and of the near-misses just below it, score 1, where the 96 missed links hide (recall). Name it: the human-in-the-loop review any automated pipeline needs before anyone trusts it.",
])
slide(26, "Lose pairs or invent them?", "35:00", [
    f"Share who changed postcode: truth {pct(MV['Truth (all true links)'])}; strict {pct(MV['Rules, strict'])}; lenient {pct(MV['Rules, lenient'])}; model {pct(MV['Rules + Haiku'])}.",
    "Postcode is a matching field and the outcome, so the linkage error is correlated with what you measure: missed links hide movers, invented links create them.",
    "No free policy. The truth is over all 5,000 true links, including the 13 blocking lost. Say that Febrl's moves are planted (that is why the truth is known) and that the pattern transfers: in their data they review random samples of the band and of the near-misses and measure both error rates, exactly like part one.",
])

section("Part D. Close (slides 27 to 30, 37:00 to 40:00)")
slide(27, "What we know works", "37:00", [
    "Three about the instrument (codebook in the prompt, structured output with fixed settings, log the instrument), two about validation (random sample by class and group, measure stability by running a sample twice), one about estimation (correct it).",
    "Credit where the points came from: definitions 14, examples nothing measurable, model upgrade 1.5.",
])
slide(28, "Report, privacy, cost (folded: thirty seconds)", "38:00", [
    "Report: codebook and prompt verbatim, model and version, validation sample, metrics by class and group, estimator and interval, where code and outputs live.",
    "Privacy: text sent to an API leaves the machine; IRB data needs a review first; a model on your own machine, a campus service or de-identification are the routes. Restricted data (an FSRDC project, a DUA): nothing leaves the enclave, you validate inside it, and disclosure review decides which cells of the error table come out. Bill titles are public; their data may not be.",
    f"Cost: 6,600 labels and 2,966 pair decisions, {((sum(v['in'] for v in R['tokens_titles'].values()) + sum(v['in'] for v in R['tokens_pairs'].values()))/1e6):.1f} million input tokens, under 20 dollars, about an hour. The 200 hand-coded titles: an afternoon that makes the dollars usable. Thirty seconds.",
])
slide(29, "Key points and take-home", "38:30", [
    "Read the five points slowly. Point 4: 200 at random is the floor for one share, a comparison costs more per group. Point 5: report the linkage policy with the number.",
    "Take-home: 1,000 documents from their project, 200 hand-coded at random, the notebook's three estimators, bring the corrected interval to a D-Lab consultation. Fewer than 200 documents: code them all.",
    "Collect the exit tickets (the Challenge 3 line on the handout).",
])
slide(30, "Thank you", "39:30", ["Stop. Smile. Wait for questions."])

section("Likely questions about the lesson, with answers")
qa = [
    ("Why bill titles and not interviews or social media, which our students actually have?",
     ["Titles are short, public, free of privacy issues, and come with the most credible gold standard available (trained coders, documented agreement). The method is what transfers; the dataset is chosen so the method can be checked against truth in front of the room.",
      "The take-home challenge moves it to their own data, and the privacy card tells them what to do before sending interviews to an API."]),
    ("Isn't 88% accuracy good enough for most purposes?",
     ["For a descriptive share across the whole corpus the naive estimate was off by one point. For the comparison between parties it created a 2.5-point gap that does not exist. On 1,650 titles that gap is marginal; on a real corpus of 20,000 documents the same bias would be six standard errors, because naive intervals shrink with N and the bias does not. Good enough depends on the comparison you make, and you only know by validating by group."]),
    ("In 1973-86 your corrected estimate lands farther from the truth than the naive one. How do you explain that to a student who just watched the correction work?",
     ["I say it before they do, on the slide. The correction is unbiased, not magic: with 63 hand-coded titles in that era the interval is about plus or minus 5 points, and the point estimate wanders inside it. The naive estimate has a smaller spread and a bias it does not report; in that era the bias happened to be small. The subset alone is worse: it found 1 health bill in 63, so its textbook interval is too narrow and misses the truth, which is the known failure of that formula near zero. The honest product is the corrected interval, and the lesson's answer is Challenge 3: a comparison across eras needs a hand-coded sample per era, more than 200 in all.",
      "It is also why I would never promise a committee a point estimate from this method; I would promise the interval and the number of hand-coded documents behind it."]),
    ("If you had to cut ten minutes live, which slides go and which exercises survive?",
     ["All three challenges survive, because the exercises are the lesson; each loses a minute. Out go the precision-recall slide (one sentence on the heatmap instead), the four-policies chart (I say the three numbers on the downstream slide), and the reporting slide (folded into the checklist of what works). The era chart stays, because it answers the opening hook.",
      "The checkpoints are at the end of Challenge 2b and the end of Challenge 3; the talking points have the cuts written at each one."]),
    ("How would you adapt this for complete beginners, or for a two-hour workshop?",
     ["Beginners: the hand-coding, the bets, the heatmap and the three-estimator chart carry the lesson without code; the notebook cells become demonstrations. Two hours: add the live API block from the APIs workshop, let pairs edit the codebook and re-score a 100-title sample, and add a second outcome (Education) to show the correction per category."]),
    ("How do you assess whether participants learned it?",
     ["Formatively, through the poll, the bets and the bell questions. The exit ticket is Challenge 3, where each person writes n per party, the hours and the interval, and it has a checkable answer: about 400 per party. For a longer course, the take-home challenge with their own data, brought to a consultation, is the real assessment."]),
    ("Your naive interval covers the truth in every panel. Where is the demonstrated bias?",
     ["The bias is 1 to 2.5 points and the naive standard error at 1,650 titles is about 1 point, so here it hides inside two standard errors; that is exactly the problem, an interval that shrinks around the wrong number. Two things show the bias is real: on the same bills labeled two ways, the model's error differs by party and flips sign after 2000, a paired comparison with z near 3; and the same 2.5-point partisan gap on 20,000 documents would sit almost six standard errors from zero. The lesson is not that naive intervals miss on 1,650 titles; it is that they will miss, with confidence, on a real corpus."]),
    ("The 200 sit inside the 1,650 and your truth is the 1,650. Are the intervals for the right target?",
     ["As implemented they target the population the 1,650 were drawn from, the usual PPI setting: the variance has a model-mean term over N and an error term over n, with no finite-population correction. Against the finite 1,650 they are conservative (in simulation they cover about 98% of the time), because the model-mean term vanishes and the error term would carry a factor of one minus n over N. The point estimate is the same as Egami and coauthors' design-based estimator; only the width differs, in the safe direction. For a student's project the population target is the one they want anyway."]),
    ("Why prediction-powered inference and not just reporting accuracy?",
     ["Accuracy describes the instrument; it does not fix the estimate. PPI and the Egami et al. estimators use the hand-coded sample to remove the bias and give a valid interval, in one line of arithmetic, whatever the error pattern. It is also the natural bridge to a methods appendix a reviewer will accept."]),
    ("What about run-to-run variability of the model?",
     ["The same title can get a different label on a second run, and the settings that used to control that are not exposed by every model. So the lesson teaches to measure stability directly: run a random sample twice and report run-to-run agreement next to accuracy; the D-Lab's Qualitative Coding with AI Tools workshop already has participants run the agent twice on one passage. In the cached run every label has a timestamp and prompt version, so the comparison is reproducible."]),
    ("The record linkage half uses synthetic data. Is the mobility rate meaningful?",
     ["No, and I say so in the room: Febrl's moves are planted, which is exactly why every true link is known and the policies can be scored. What transfers is the structure: rules for the easy pairs, a reviewed random sample of the band, precision and recall by group, and a downstream number that moves with the policy. The full Record Linkage workshop would use a real pair of administrative files with a clerical review sample, which is the D-Lab's consulting case."]),
    ("A student has 800 interview transcripts under an IRB protocol that forbids third-party processing. Walk us from the model she may use to what disclosure review does to her error table.",
     ["First the model: nothing goes to a commercial API; a model on her own machine or a campus-hosted service, and de-identification is not a free pass, because transcripts re-identify people through details, not names. Then the same recipe inside the environment: a random hand-coded sample, accuracy and precision and recall by the groups she compares, the corrected estimate. Then disclosure: the confusion matrix is a table of small cells, and a cell with three people in it is a disclosure risk, so the error table she publishes is the aggregated version the review allows, and the protocol or the DUA is amended before any of it leaves. The order matters: data use review, validation inside, disclosure review, publication."]),
    ("Two Data Science Fellows will teach this next semester. What do you hand them, how do you know they are ready, what changes after the first run?",
     ["They get the repository on the D-Lab template: the notebooks with the challenges and solutions, the slides with the speaker notes, the handout, and a short instructor note with the three bets as pacing checkpoints. Ready means one rehearsal in front of me with the clock running and their own answers to the Challenge 2 disputes. After the first run I look at two things: the exit tickets (the Challenge 3 numbers tell me whether the curve was understood) and the questions the room asked. Two changes are already planned: a model-deprecation plan, because the cached labels name a model that will not exist forever, and a challenge each fellow adds from their own field, which is how the lesson stops being mine."]),
    ("How does this connect to D-Lab consulting and the FSRDC?",
     ["Most 'can I use ChatGPT on my data' consultations are questions about validation, privacy and cost; the lesson gives a protocol and a checklist. For restricted data, the linkage half is the vocabulary a researcher needs before a proposal: blocking, the ambiguous band, precision and recall by group, and what a reviewed sample costs."]),
    ("Why two models? Why these two?",
     ["To separate the instrument from the model as sources of error. The result is that the prompt moves accuracy by 14 points and the model upgrade by 1.5, which is itself a teaching point about where to spend effort. A cheap and a strong model from one provider keeps the comparison clean and the cost low."]),
    ("The objectives say participants will design a validation sample and read a corrected estimate. Do they actually do that in 45 minutes?",
     ["In the room they do the parts that need judgment: they code by hand and feel the ceiling, they predict and read the metrics, they vote on whose error it is, and they read the curve and choose n. The notebook does the arithmetic, and the take-home asks them to run the three estimators on their own data. The verbs on the objectives match that split: explain, design, explain, read, apply."]),
    ("How would this work in R?",
     ["The notebook's logic is a few data frames and one estimator; an R twin (tidyverse plus the same cached CSVs) is a day of work and matches the D-Lab practice of offering both languages."]),
]
for q, ans in qa:
    heading(doc, q, size=11, before=6, after=1)
    for a in ans:
        bullet(doc, a, after=2)
doc.save(PROJECT / "docs/Talking_Points_and_QA.docx")
print("talking points written")
