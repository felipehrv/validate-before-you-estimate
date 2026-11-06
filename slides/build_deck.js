// Build the lesson deck: Validate Before You Estimate (D-Lab mock lesson)
// Run: node slides/build_deck.js
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const sharp = require("sharp");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const ROOT = path.resolve(__dirname, "..");
const R = JSON.parse(fs.readFileSync(path.join(ROOT, "outputs/results.json")));
const X = JSON.parse(fs.readFileSync(path.join(ROOT, "outputs/exercise_items.json")));
const FIG = (n) => path.join(ROOT, "figures", n);

const THEME = {
  name: "Validate Before You Estimate",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B1F23", lt1: "FFFFFF", dk2: "2F3A44", lt2: "F3F4F2",
    accent1: "2A78D6", accent2: "EB6834", accent3: "1BAF7A",
    accent4: "EDA100", accent5: "6250D6", accent6: "898781",
    hlink: "2A78D6", folHlink: "6250D6",
  },
};
const H = THEME.colors; // hex, for the few options that take hex only

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "Felipe H. Rodriguez Villalta";
pres.title = "Validate Before You Estimate";
const C = pres.SchemeColor;

const W = 13.33, HGT = 7.5, M = 0.6;
const FOOT = "Validate Before You Estimate  |  Mock lesson, D-Lab, UC Berkeley";

// ---------------------------------------------------------------- layouts
pres.defineSlideMaster({
  title: "DARK", background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 2.2, w: W - 2 * M, h: 1.6, fontSize: 44, bold: true, color: C.background1, valign: "bottom", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: M, y: 3.95, w: W - 2 * M, h: 1.2, fontSize: 22, color: "CADCFC", valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "LIGHT", background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.38, w: W - 2 * M, h: 0.9, fontSize: 30, bold: true, color: C.text1, valign: "middle", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: M, y: HGT - 0.45, w: 9, h: 0.3, fontSize: 10, color: C.accent6, margin: 0 } } },
  ],
  slideNumber: { x: W - 1.2, y: HGT - 0.45, w: 0.6, h: 0.3, fontSize: 10, color: H.accent6 },
});
pres.defineSlideMaster({
  title: "TINT", background: { color: C.background2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.38, w: W - 2 * M, h: 0.9, fontSize: 30, bold: true, color: C.text1, valign: "middle", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: M, y: HGT - 0.45, w: 9, h: 0.3, fontSize: 10, color: C.accent6, margin: 0 } } },
  ],
  slideNumber: { x: W - 1.2, y: HGT - 0.45, w: 0.6, h: 0.3, fontSize: 10, color: H.accent6 },
});

// ---------------------------------------------------------------- helpers
const dims = {};
async function fig(slide, name, box, opts = {}) {
  const p = FIG(name);
  if (!dims[name]) { const m = await sharp(p).metadata(); dims[name] = m.width / m.height; }
  const ar = dims[name];
  let w = box.w, h = w / ar;
  if (h > box.h) { h = box.h; w = h * ar; }
  const x = box.x + (box.w - w) / 2, y = box.y + (opts.top ? 0 : (box.h - h) / 2);
  slide.addImage({ path: p, x, y, w, h, objectName: opts.alt || name });
}
function bullets(slide, items, box, opts = {}) {
  const fs_ = opts.fontSize || 18;
  const arr = items.map((t, i) => {
    const o = { bullet: opts.numbered ? { type: "number" } : true, fontSize: fs_, color: C.text1, paraSpaceAfter: opts.space || 8, breakLine: i < items.length - 1 };
    if (typeof t === "string") return { text: t, options: o };
    return { text: t.text, options: { ...o, bold: !!t.bold } };
  });
  slide.addText(arr, { x: box.x, y: box.y, w: box.w, h: box.h, valign: "top", margin: 0, isTextBox: true, objectName: opts.name || "bullets" });
}
function label(slide, text, x, y, w, color, opts = {}) {
  slide.addText(text, { x, y, w, h: 0.32, fontSize: opts.fontSize || 12, bold: true, color, margin: 0, isTextBox: true, charSpacing: 2, objectName: opts.name || "label" });
}
function para(slide, text, box, opts = {}) {
  slide.addText(text, { x: box.x, y: box.y, w: box.w, h: box.h, fontSize: opts.fontSize || 18, color: opts.color || C.text1, valign: opts.valign || "top", margin: 0, isTextBox: true, italic: !!opts.italic, bold: !!opts.bold, align: opts.align || "left", objectName: opts.name || "paragraph" });
}
function numCircle(slide, n, x, y, color) {
  slide.addShape(pres.ShapeType.ellipse, { x, y, w: 0.5, h: 0.5, fill: { color }, line: { color, width: 0 }, objectName: "circle" });
  slide.addText(String(n), { x, y, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "circle-number" });
}
function card(slide, box, title, body, color, opts = {}) {
  slide.addShape(pres.ShapeType.roundRect, { x: box.x, y: box.y, w: box.w, h: box.h, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: "E1E0D9", width: 0.75 }, objectName: "card" });
  slide.addShape(pres.ShapeType.ellipse, { x: box.x + 0.2, y: box.y + 0.2, w: 0.34, h: 0.34, fill: { color }, line: { color, width: 0 }, objectName: "card-dot" });
  slide.addText(title, { x: box.x + 0.65, y: box.y + 0.14, w: box.w - 0.8, h: 0.46, fontSize: opts.titleSize || 16, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "card-title" });
  slide.addText(body, { x: box.x + 0.2, y: box.y + 0.72, w: box.w - 0.4, h: box.h - 0.85, fontSize: opts.bodySize || 13, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "card-body" });
}
const pct = (v, d = 1) => (100 * v).toFixed(d) + "%";
const pts = (v, d = 1) => (100 * v).toFixed(d);
const E = R.estimates;
const S = Object.fromEntries(R.showdown.map((r) => [r.model + "|" + r.prompt, r]));
const haikuP1 = S["claude-haiku-5-5|prompt_1_names"], haikuP2 = S["claude-haiku-5-5|prompt_2_codebook"];
const haikuP3 = S["claude-haiku-5-5|prompt_3_codebook_examples"], sonnetP3 = S["claude-sonnet-5-5|prompt_3_codebook_examples"];
const BP = R.by_party_haiku_p3, BE = R.by_era_haiku_p3;
const LP = R.linkage_policies, MV = R.moved_rate, BL = R.blocking;
const PJ = R.sample_size_curve_projected, PJrow = Object.fromEntries(PJ.rows.map((r) => [r.n_hand_coded, r]));
// the naive partisan gap on a corpus of 20,000 titles: how many standard errors it would look like
const nD20 = 20000 * BP.Democrat.n / (BP.Democrat.n + BP.Republican.n), nR20 = 20000 - nD20;
const seGap20 = Math.sqrt(E.Democrat.naive.est * (1 - E.Democrat.naive.est) / nD20 + E.Republican.naive.est * (1 - E.Republican.naive.est) / nR20);
const zGap20 = (E.Democrat.naive.est - E.Republican.naive.est) / seGap20;
const eraHW = Math.max(...["1973-1986", "1987-2000", "2001-2016"].map((e) => 1.96 * E[e].corrected.se));
const eraHWmin = Math.min(...["1973-1986", "1987-2000", "2001-2016"].map((e) => 1.96 * E[e].corrected.se));
const nC3 = Math.ceil(PJ.n_for_target_corrected / 50) * 50, nH3 = Math.ceil(PJ.n_for_target_human_only / 50) * 50;

(async () => {
  let s;
  // ================================================================ OVERVIEW
  pres.addSection({ title: "Overview" });

  // 1. Title
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Overview" });
  s.addText("Validate Before You Estimate", { placeholder: "title" });
  s.addText("Using AI to label text and link records in social science research", { placeholder: "body" });
  para(s, "Felipe H. Rodriguez Villalta  |  Mock lesson for the D-Lab, UC Berkeley  |  November 5, 2026", { x: M, y: 6.3, w: W - 2 * M, h: 0.4 }, { fontSize: 14, color: "CADCFC" });
  para(s, "A 45-minute hands-on lesson for graduate students", { x: M, y: 1.5, w: W - 2 * M, h: 0.5 }, { fontSize: 16, color: "CADCFC" });
  s.addNotes("Open with one line about who you are: a research scientist who has spent ten years making other people's research data usable, and who teaches the way he works. Then: this is a 45-minute lesson; today you will see the overview first, then I will teach most of it with you as my students.");

  // 2. Why this lesson
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Overview" });
  s.addText("Why this lesson, and why now", { placeholder: "title" });
  card(s, { x: M, y: 1.5, w: 3.9, h: 2.5 }, "The practice", "Graduate students already use large language models (LLMs) to code interviews, open-ended answers, documents and posts. It is cheap, fast and often good.", C.accent1);
  card(s, { x: M + 4.1, y: 1.5, w: 3.9, h: 2.5 }, "The gap", "Desai, Card and Jacobs (2026) reviewed 27 papers in eight flagship social science journals: LLM measurements sat at the center of the analysis, and validation was limited and inconsistent.", C.accent2);
  card(s, { x: M + 8.2, y: 1.5, w: 3.9, h: 2.5 }, "The risk", "Egami and coauthors show that treating predicted labels as if they were true biases estimates and breaks confidence intervals even when accuracy is above 90%.", C.accent5);
  para(s, "The one idea the lesson teaches", { x: M, y: 4.4, w: 6, h: 0.4 }, { fontSize: 13, color: C.accent6, bold: true });
  para(s, "Every label an AI gives you is a measurement with error. Measure the error against people, then carry it into your estimate.", { x: M, y: 4.8, w: W - 2 * M, h: 1.3 }, { fontSize: 24, bold: true });
  s.addNotes("Three cards: the practice is common, the validation is thin, and the consequence is statistical, not cosmetic. Then the one-sentence idea. Everything else in the lesson is this sentence applied twice: to text labels and to record links.");

  // 3. Learning objectives
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Overview" });
  s.addText("Learning objectives", { placeholder: "title" });
  para(s, "After this lesson, participants will be able to:", { x: M, y: 1.4, w: 8, h: 0.4 }, { fontSize: 15, color: C.accent6 });
  const objs = [
    "Explain why a classification prompt is a measurement instrument, and what changes when the codebook goes into it.",
    "Design a validation sample and read agreement metrics (accuracy, kappa, precision and recall) as measurement error.",
    "Explain why high accuracy does not guarantee an unbiased estimate or a valid comparison between groups.",
    "Read a bias-corrected estimate that combines AI labels with a small hand-coded sample, and choose how many documents to code by hand from an interval curve.",
    "Apply the same recipe to record linkage: rules for the easy pairs, an AI adjudicator for the ambiguous band, a reviewed sample to measure the error.",
  ];
  objs.forEach((t, i) => {
    numCircle(s, i + 1, M, 1.95 + i * 0.95, i < 4 ? H.accent1 : H.accent3);
    para(s, t, { x: M + 0.7, y: 1.95 + i * 0.95, w: W - 2 * M - 0.7, h: 0.85 }, { fontSize: 17, valign: "middle" });
  });
  s.addNotes("Five objectives, each one checked by an exercise. Note the verbs: explain, design, explain, read, apply. In the notebook, read becomes run: the take-home asks them to compute the three estimators on their own data. Objective 3 is the heart of the lesson and the one most AI workshops skip.");

  // 4. Lesson at a glance: timeline + exercise goals
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Overview" });
  s.addText("The lesson at a glance: 45 minutes, four exercises", { placeholder: "title" });
  const blocks = [
    ["0-5", "Hook and poll", C.accent6], ["5-13", "Codebook to prompt", C.accent1], ["13-24", "Validate", C.accent1],
    ["24-32", "Estimate", C.accent1], ["32-40", "Link records", C.accent3], ["40-45", "What works, report, wrap", C.accent6],
  ];
  const mins = [5, 8, 11, 8, 8, 5];
  let xx = M; const total = 45, tw = W - 2 * M;
  blocks.forEach((b, i) => {
    const bw = (mins[i] / total) * tw;
    s.addShape(pres.ShapeType.rect, { x: xx, y: 1.5, w: bw - 0.04, h: 0.75, fill: { color: b[2] }, line: { color: b[2], width: 0 }, objectName: "timeline-block" });
    s.addText(b[1], { x: xx, y: 1.5, w: bw - 0.04, h: 0.75, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0.03, isTextBox: true, objectName: "timeline-label" });
    s.addText(b[0] + " min", { x: xx, y: 2.28, w: bw, h: 0.3, fontSize: 10, color: C.accent6, align: "center", margin: 0, isTextBox: true, objectName: "timeline-minutes" });
    xx += bw;
  });
  const rows = [
    [{ text: "Exercise", options: { bold: true, color: H.lt1, fill: { color: H.dk2 } } }, { text: "What participants do", options: { bold: true, color: H.lt1, fill: { color: H.dk2 } } }, { text: "Goal (what it lets them do)", options: { bold: true, color: H.lt1, fill: { color: H.dk2 } } }],
    ["Challenge 1  (min 8)", "Code six bill titles by hand on the handout; compare with a neighbor", "Feel the codebook's ambiguity; set the human ceiling the model is measured against"],
    ["Challenge 2  (min 14)", "Predict which of three prompts wins; then adjudicate three disputed titles", "Separate prompt from model as sources of error; decide whose error it is"],
    ["Challenge 3  (min 28)", "Read the interval curve; choose how many titles to hand-code per group to see a 3-point gap", "Make the design decision every project faces, with a checkable answer"],
    ["Question  (min 37)", "Lose pairs or invent them? Decide for the number you want to know", "See that linkage policy is an estimation choice"],
  ];
  s.addTable(rows, { x: M, y: 2.85, w: W - 2 * M, colW: [2.2, 5.2, 4.73], fontSize: 12.5, color: H.dk1, border: { type: "solid", color: "E1E0D9", pt: 0.75 }, rowH: 0.62, valign: "middle", margin: 0.06, objectName: "exercise-table" });
  s.addNotes("An exercise every ten to twelve minutes; the D-Lab guide asks for at least one every fifteen. Each exercise has a goal stated as something the participant can do afterward, not as content covered.");

  // 5. Teaching approach
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Overview" });
  s.addText("Teaching approach", { placeholder: "title" });
  const appr = [
    ["Backward design", "Started from the decisions a graduate student has to make in their own project (how many to hand-code, what to report) and built the exercises to practice those decisions.", C.accent1],
    ["Predict, observe, explain", "Before any result is shown, participants commit to a guess: which prompt wins, which group the model misreads. The gap between the guess and the data is what sticks.", C.accent2],
    ["Hands-on without a key", "Every call to the model was run beforehand and cached, so the notebook runs on DataHub or Binder with no API key, or on a laptop with no internet at all. A saved dry-run file shows exactly what one call sends and returns.", C.accent3],
    ["A path without code", "Hand-coding, voting and error analysis need no Python. A participant from history or anthropology can take the full lesson and leave with the method.", C.accent5],
    ["Evaluation and ethics inside the method", "Validation, differential error, privacy of data sent to an API, and cost are steps of the workflow, not a closing slide.", C.accent4],
    ["Formative assessment", "An opening poll, bell questions as attention checks, and Challenge 3 as a measurable exit ticket.", C.accent6],
  ];
  appr.forEach((a, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    card(s, { x: M + col * 6.1, y: 1.45 + row * 1.85, w: 6.0, h: 1.72 }, a[0], a[1], a[2], { bodySize: 12.5, titleSize: 15 });
  });
  s.addNotes("Six design choices. Say them as choices, with the reason. The D-Lab's own guide asks for an exercise every fifteen minutes, colloquial language, and no caveats on slides; the design follows it.");

  // 6. Where it sits in the catalog
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Overview" });
  s.addText("Where it sits in the D-Lab catalog", { placeholder: "title" });
  const boxes = [
    ["Python APIs for Large Language Models", "How to call a model and get structured output", M, 1.6, C.accent6],
    ["Qualitative Coding with AI Tools", "How to audit an agent's codes against the source text", M, 3.0, C.accent6],
    ["This lesson: Validate Before You Estimate", "How to measure the error of AI labels and links, and carry it into the estimate", M + 4.15, 2.3, C.accent1],
    ["Next workshop: Record Linkage for Research", "Blocking, comparison, probabilistic matching, clerical review and privacy, with AI as one adjudicator", M + 8.3, 2.3, C.accent3],
  ];
  boxes.forEach((b) => {
    s.addShape(pres.ShapeType.roundRect, { x: b[2], y: b[3], w: 3.5, h: 1.5, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: b[4] === C.accent6 ? "E1E0D9" : H.accent1, width: b[4] === C.accent6 ? 0.75 : 2 }, objectName: "catalog-box" });
    s.addText(b[0], { x: b[2] + 0.16, y: b[3] + 0.12, w: 3.2, h: 0.55, fontSize: 13.5, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "catalog-title" });
    s.addText(b[1], { x: b[2] + 0.16, y: b[3] + 0.7, w: 3.2, h: 0.75, fontSize: 11.5, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "catalog-body" });
  });
  // arrows
  [[M + 3.5, 2.35, M + 4.15, 3.05], [M + 3.5, 3.75, M + 4.15, 3.05], [M + 7.65, 3.05, M + 8.3, 3.05]].forEach((a) => {
    s.addShape(pres.ShapeType.line, { x: a[0], y: a[1], w: a[2] - a[0], h: a[3] - a[1], line: { color: H.accent6, width: 1.5, endArrowType: "triangle" }, flipV: a[3] < a[1], objectName: "arrow" });
  });
  bullets(s, [
    "What is new here: the method comes from Berkeley (prediction-powered inference) and Princeton (design-based supervised learning); this lesson is a 45-minute, no-code version with a hand-coding budget decision, and the same recipe applied to record linkage.",
    "Builds on Python APIs for Large Language Models and Prompt Engineering; Python Fundamentals recommended to run the notebook afterward.",
    "Format: two Jupyter notebooks (lessons and solutions), a handout, cached outputs; fits as a 50-minute AI Pulse session or as part 3 of a two-hour workshop.",
  ], { x: M, y: 4.6, w: W - 2 * M, h: 2.1 }, { fontSize: 13, space: 7 });
  s.addNotes("The catalog already teaches the mechanics (Python APIs for LLMs, with structured output through OpenRouter) and the qualitative audit (Qualitative Coding with AI Tools: code six passages by hand, run the agent twice, audit the codes against the text). This lesson is the measurement step between them, and the natural next workshop is record linkage in full, which the posting asks for.");

  // 7. Materials
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Overview" });
  s.addText("Materials and what participants need", { placeholder: "title" });
  card(s, { x: M, y: 1.5, w: 5.9, h: 2.3 }, "Data", "Comparative Agendas Project, U.S. congressional bills: 1,650 titles, 75 per Congress from 1973 to 2016, each coded by trained human coders (90 to 95% agreement on major topic). Febrl: two files of 5,000 person records with the 5,000 true links known.", C.accent1, { bodySize: 12.5 });
  card(s, { x: M + 6.2, y: 1.5, w: 5.9, h: 2.3 }, "AI workflow", "Three prompts for the same codebook, two models, 6,600 labels and 2,966 pair decisions, all run beforehand and cached with model name, prompt version, tokens and timestamp. Cost of the whole run: under 20 dollars.", C.accent2, { bodySize: 12.5 });
  card(s, { x: M, y: 4.0, w: 5.9, h: 2.3 }, "Repository", "D-Lab template: README with objectives and prerequisites, lessons and solutions notebooks with the challenge, question, tip and warning conventions, data folder, prompts folder, dry-run output. Runs on DataHub and Binder, so a Data Science Fellow can teach it after one rehearsal.", C.accent3, { bodySize: 12.5 });
  card(s, { x: M + 6.2, y: 4.0, w: 5.9, h: 2.3 }, "In the room", "A two-page handout for the hand-coding, voting and budget exercises; slides. Nothing to install, no account to create.", C.accent5, { bodySize: 12.5 });
  s.addNotes("Keep this to thirty seconds. The point is that the lesson is reproducible and that it runs without anyone needing an API key.");

  // 8. Transition
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Overview" });
  s.addText("For the next thirty minutes, you are my graduate students", { placeholder: "title" });
  s.addText("Three questions you will be able to answer when we are done:\n1. Can I trust this label?   2. Can I trust this estimate?   3. How many do I have to code by hand?", { placeholder: "body" });
  para(s, "It is OK not to know. Every bet in this lesson is a guess, and the wrong guesses are the ones that teach.", { x: M, y: 5.45, w: W - 2 * M, h: 0.5 }, { fontSize: 18, color: "F9A27B", italic: true });
  s.addNotes("Switch register here. Slow down, make eye contact, hand out the sheet. Say the IOKN2K line out loud; it is the D-Lab's own. From now on you teach.");

  // ================================================================ PART 1: TEXT
  pres.addSection({ title: "Part 1: Labels as measurements" });

  // 9. Hook
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Part 1: Labels as measurements" });
  s.addText("Would you put this chart in your dissertation?", { placeholder: "title" });
  await fig(s, "f00_hook_naive_series.png", { x: M, y: 1.4, w: 7.6, h: 5.2 }, { alt: "Naive series of the health share by era" });
  para(s, "Quick poll, hands up", { x: 8.6, y: 1.6, w: 4.2, h: 0.4 }, { fontSize: 13, color: C.accent6, bold: true });
  bullets(s, ["Have you used an LLM to code text for research?", "Did you check it against a person?", "Did the check change a number in your results?"], { x: 8.6, y: 2.05, w: 4.2, h: 2.6 }, { fontSize: 17, space: 12 });
  para(s, "The chart is real. The labels came from an LLM with 88% accuracy. By the end of the lesson you will know how much of that climb is the bills and how much is the instrument.", { x: 8.6, y: 4.8, w: 4.2, h: 1.6 }, { fontSize: 15, italic: true });
  s.addNotes("Let the poll run. Most hands go up on the first question and down on the third. That gap is the lesson. Do not say yet what is wrong with the chart; slide 22 answers it: the rise is real, the steady climb after 2000 is the instrument.");

  // 10. Data
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Part 1: Labels as measurements" });
  s.addText("The data: bill titles with a human answer key", { placeholder: "title" });
  await fig(s, "f01_sample_composition.png", { x: M, y: 1.4, w: 7.4, h: 5.2 }, { alt: "Sample composition by label" });
  bullets(s, [
    "Comparative Agendas Project (CAP): every U.S. bill since 1947, coded into 21 policy topics by trained coders.",
    "We keep seven labels: Health, Education, Environment, Immigration, Law and Crime, Social Welfare, and Other.",
    "1,650 titles, 75 per Congress, 1973 to 2016, so every share in this lesson is the average Congress's share. Sponsor's party and chamber come with each bill.",
    "The human codes are our gold standard. Trained coders agree 90 to 95% of the time, so even gold has error. Keep that in mind.",
    "Your version: interview transcripts coded for housing insecurity by respondent gender; nineteenth-century newspaper ads coded for occupation by decade; clinic notes coded for tobacco use by site.",
  ], { x: 8.3, y: 1.5, w: 4.5, h: 5.0 }, { fontSize: 13.5, space: 8 });
  s.addNotes("Why bill titles: short, public, no privacy issue, and a credible gold standard. Say that the 'Other' bucket is 72% of bills, which already tells you accuracy alone will flatter the model.");

  // 11. Codebook to prompt
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Part 1: Labels as measurements" });
  s.addText("From codebook to prompt: three instruments for one task", { placeholder: "title" });
  const inst = [
    ["P1  Category names only", "\"Classify this title into one of: Health, Education, ... Other.\" Nothing else. The model fills the definitions from its own head.", C.accent6],
    ["P2  The codebook", "The same, plus a condensed version of the CAP definitions for each topic, including the boundary rules (drug treatment is Health; drug trafficking is Law and Crime; veterans' programs are Other).", C.accent1],
    ["P3  Codebook plus examples", "P2 plus fourteen titles coded by humans, two per topic, drawn from outside the evaluation sample so nothing in the prompt is ever scored.", C.accent2],
  ];
  inst.forEach((a, i) => card(s, { x: M + i * 4.1, y: 1.45, w: 3.95, h: 2.55 }, a[0], a[1], a[2], { bodySize: 12.5, titleSize: 14.5 }));
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 4.25, w: W - 2 * M, h: 2.0, rectRadius: 0.08, fill: { color: "1B1F23" }, line: { color: "1B1F23", width: 0 }, objectName: "code-box" });
  s.addText([
    { text: "Every call, every prompt, every title:  ", options: { color: "CADCFC", fontSize: 13 } },
    { text: "{\"label\": \"<topic>\", \"why\": \"<one short sentence>\"}", options: { color: "F0EFEC", fontSize: 15, fontFace: "Courier New", breakLine: true } },
    { text: "Logged with the label: model name, prompt version, tokens in and out, timestamp. The prompt is frozen before anything is scored.", options: { color: "CADCFC", fontSize: 13, breakLine: true } },
    { text: "Plain words: JSON is a fixed answer format the computer can read. Tokens are the word pieces the model reads and bills you for. Haiku and Sonnet are a small and a large model of the same family. Reasoning off means the model answers directly, without a hidden think-aloud.", options: { color: "898781", fontSize: 11.5 } },
  ], { x: M + 0.3, y: 4.4, w: W - 2 * M - 0.6, h: 1.7, valign: "middle", margin: 0, isTextBox: true, objectName: "code-text" });
  s.addNotes("The word instrument is deliberate. A prompt is a questionnaire for the model; a changed prompt is a changed instrument and gets its own validation. Nothing in a prompt is ever scored.");

  // 12. Challenge 1
  s = pres.addSlide({ masterName: "TINT", sectionTitle: "Part 1: Labels as measurements" });
  s.addText("Challenge 1: be the coder", { placeholder: "title" });
  numCircle(s, 1, M, 1.45, H.accent1);
  para(s, "On the handout: six real bill titles and the seven-topic codebook. Code each one. Three minutes. Then compare with your neighbor.", { x: M + 0.7, y: 1.4, w: 11.5, h: 0.7 }, { fontSize: 18, valign: "middle" });
  X.hand_coding.forEach((h, i) => {
    numCircle(s, i + 1, M, 2.35 + i * 0.68, H.accent6);
    para(s, `(${h.year}) ${h.title}`, { x: M + 0.7, y: 2.33 + i * 0.68, w: 11.5, h: 0.62 }, { fontSize: 12.5, valign: "middle" });
  });
  para(s, "Bell question: on which one did you and your neighbor disagree, and what rule would have settled it?", { x: M, y: 6.5, w: 11.5, h: 0.4 }, { fontSize: 14, italic: true, color: C.accent6 });
  s.addNotes("Give them three real minutes. Walk around. Ask two pairs where they disagreed before showing the results. The disagreement is the point: the codebook has boundaries that reasonable people draw differently.");

  // 13. Reveal of Challenge 1
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Part 1: Labels as measurements" });
  s.addText("What the trained coders said, and what it means", { placeholder: "title" });
  const revealRows = [[{ text: "Title", options: { bold: true, color: H.lt1, fill: { color: H.dk2 } } }, { text: "Human code", options: { bold: true, color: H.lt1, fill: { color: H.dk2 } } }, { text: "The tempting alternative", options: { bold: true, color: H.lt1, fill: { color: H.dk2 } } }]];
  const alt = { Health: "Law and Crime (drug enforcement)", "Law and Crime": "Social Welfare (a pension)", "Social Welfare": "Other (a tax provision)", Education: "Other (a tax bond)", Immigration: "Law and Crime (enforcement)", Other: "Environment (parks and land)" };
  X.hand_coding.forEach((h) => revealRows.push([h.title.length > 175 ? h.title.slice(0, h.title.lastIndexOf(" ", 172)) + " ..." : h.title, h.gold, alt[h.gold]]));
  s.addTable(revealRows, { x: M, y: 1.35, w: W - 2 * M, colW: [7.6, 1.8, 2.73], fontSize: 10, color: H.dk1, border: { type: "solid", color: "E1E0D9", pt: 0.75 }, rowH: 0.44, valign: "middle", margin: 0.03, objectName: "reveal-table" });
  bullets(s, [
    "Where you disagreed with your neighbor is where the model will disagree with the humans. Human agreement is the ceiling, not 100%.",
    "Titles 3 and 4 change the tax law, and the codebook lists taxes under Other. The CAP coders put them under Social Welfare and Education, following a rule the condensed codebook states in one sentence: code the bill by who it is for (the older workers, the schools), not by the tool it uses. If you missed it, so did the model: many of its errors are this kind of slip into Other.",
    "The answer key is a measurement too. The CAP coders agree 90 to 95% of the time, and a real case from the CAP data: a bill funding treatment centers for torture victims is coded Environment. Keep that in mind in Challenge 2: sometimes the human is the one who slipped.",
  ], { x: M, y: 4.55, w: W - 2 * M, h: 2.35 }, { fontSize: 13, space: 6 });
  s.addNotes("Do not rush these results. Two messages: the ceiling, and the fact that gold has error. Both come back in the estimation block.");

  // 14. Challenge 2a: predict
  s = pres.addSlide({ masterName: "TINT", sectionTitle: "Part 1: Labels as measurements" });
  s.addText("Challenge 2, part one: place your bet", { placeholder: "title" });
  numCircle(s, 2, M, 1.45, H.accent1);
  para(s, "Same 1,650 titles, same cheap model, three prompts. Which instrument agrees most with the human coders, and by how much? Mark your guess on the handout.", { x: M + 0.7, y: 1.4, w: 11.5, h: 0.8 }, { fontSize: 18, valign: "middle" });
  const bets = [["P1", "Category names only", C.accent6], ["P2", "The codebook", C.accent1], ["P3", "Codebook plus examples", C.accent2]];
  bets.forEach((b, i) => {
    s.addShape(pres.ShapeType.roundRect, { x: M + i * 4.1, y: 2.6, w: 3.95, h: 2.2, rectRadius: 0.1, fill: { color: C.background1 }, line: { color: "E1E0D9", width: 1 }, objectName: "bet-card" });
    s.addText(b[0], { x: M + i * 4.1, y: 2.75, w: 3.95, h: 1.0, fontSize: 44, bold: true, color: b[2], align: "center", margin: 0, isTextBox: true, objectName: "bet-code" });
    s.addText(b[1], { x: M + i * 4.1, y: 3.8, w: 3.95, h: 0.6, fontSize: 16, color: C.text1, align: "center", margin: 0, isTextBox: true, objectName: "bet-name" });
  });
  bullets(s, ["Second bet: does a stronger model (Sonnet instead of Haiku) with P3 beat the best Haiku prompt? By more or less than five points?", "Third bet, for the brave: what accuracy would a lazy coder who answers \"Other\" every time get?"], { x: M, y: 5.1, w: W - 2 * M, h: 1.4 }, { fontSize: 15, space: 10 });
  s.addNotes("Make them commit with a show of hands per prompt. The third bet sets up the kappa discussion: 'Other' is 72% of the sample, so always-Other scores 72% accuracy and a kappa of zero.");

  // 15. Reveal showdown
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Part 1: Labels as measurements" });
  s.addText("The results: the prompt moves the needle, the model barely does", { placeholder: "title" });
  await fig(s, "f02_prompt_showdown.png", { x: M, y: 1.35, w: 8.6, h: 5.3 }, { alt: "Accuracy and kappa by instrument" });
  bullets(s, [
    { text: `Names only: ${pct(haikuP1.accuracy)}. Codebook: ${pct(haikuP2.accuracy)}. Codebook plus examples: ${pct(haikuP3.accuracy)}.`, bold: true },
    `The definitions added ${pts(haikuP2.accuracy - haikuP1.accuracy, 0)} points; the examples, ${pts(haikuP3.accuracy - haikuP2.accuracy)}, which is inside run-to-run noise. A stronger model (Sonnet) adds ${pts(sonnetP3.accuracy - haikuP3.accuracy)} points, and that one is real.`,
    `A lazy coder who always answers "Other" scores 72% and a kappa of 0.00. Kappa asks how much better than chance agreement the instrument is, from 0 (chance) to 1 (perfect); P3 scores ${haikuP3.kappa.toFixed(2)}.`,
    "The instrument is the thing you control, and the thing you must report.",
  ], { x: 9.4, y: 1.5, w: 3.4, h: 5.0 }, { fontSize: 13.5, space: 10 });
  s.addNotes("Pause on the kappa point: accuracy flatters any classifier when one class dominates. Then the message: the codebook in the prompt is worth fourteen points; the examples, nothing you can measure (0.3 points; a paired test on the 1,650 titles cannot tell P2 from P3); the model upgrade, one and a half, and that one is real.");

  // 16. Confusion heatmap
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Part 1: Labels as measurements" });
  s.addText("One number hides seven stories: where the disagreements live", { placeholder: "title" });
  await fig(s, "f03_confusion_heatmap.png", { x: M, y: 1.3, w: 7.2, h: 5.5 }, { alt: "Confusion matrix heatmap" });
  const pcH = Object.fromEntries(R.per_class_haiku_p3.map((r) => [r.label, r]));
  para(s, "How to read it", { x: 8.0, y: 1.5, w: 4.8, h: 0.4 }, { fontSize: 13, color: C.accent6, bold: true });
  bullets(s, [
    "Each row is what the human coders said; each column is what the model said. The diagonal is where they agree.",
    `Health row: ${R.confusion_haiku_p3.Health.Health} health bills the model also called Health; ${R.confusion_haiku_p3.Other.Health} that it called Other. Other row: ${R.confusion_haiku_p3.Health.Other} bills the humans coded Other that the model called Health.`,
    "Most disagreements involve Other, in both directions. Other is the catch-all category, and the line between Other and a specific topic is the hardest one to draw, in this codebook and in yours.",
    `Small categories pay for the big one's accuracy: Social Welfare keeps ${pct(pcH["Social Welfare"].recall, 0)} of its bills, Education ${pct(pcH.Education.recall, 0)}.`,
  ], { x: 8.0, y: 1.95, w: 4.8, h: 4.8 }, { fontSize: 13.5, space: 10 });
  s.addNotes("Read the heatmap row by row with them. Rows are the humans. Point at the Health row and the Other column, then the Other row and the Health column. Those two cells are the whole story of the next slides.");

  // 16b. Precision and recall by class
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Part 1: Labels as measurements" });
  s.addText("Precision and recall: two questions for every category", { placeholder: "title" });
  await fig(s, "f04_precision_recall_by_class.png", { x: M, y: 1.35, w: 7.6, h: 5.2 }, { alt: "Precision and recall by class" });
  card(s, { x: 8.5, y: 1.5, w: 4.3, h: 1.75 }, "Precision", `When the model says Health, how often the humans agree. Health: ${pct(pcH.Health.precision, 0)}. One in five "Health" labels is something else.`, C.accent1, { bodySize: 12.5 });
  card(s, { x: 8.5, y: 3.4, w: 4.3, h: 1.75 }, "Recall", `When the humans say Health, how often the model notices. Health: ${pct(pcH.Health.recall, 0)}. It misses one health bill in eight.`, C.accent2, { bodySize: 12.5 });
  para(s, "A reviewer will ask for these by class, not for one accuracy number. The two questions trade off: a looser definition raises recall and lowers precision.", { x: 8.5, y: 5.35, w: 4.3, h: 1.3 }, { fontSize: 12.5, italic: true });
  s.addNotes("Two questions, two words. Say them with the Health example and move on; the handout repeats them in the linkage question.");

  // 17. Challenge 2b: whose error
  s = pres.addSlide({ masterName: "TINT", sectionTitle: "Part 1: Labels as measurements" });
  s.addText("Challenge 2, part two: whose error is it?", { placeholder: "title" });
  numCircle(s, 2, M, 1.4, H.accent1);
  para(s, "Three titles where the model and the human coder disagree. For each, vote: the model is wrong, the human is wrong, or the codebook is ambiguous. Then say what you would change.", { x: M + 0.7, y: 1.35, w: 11.5, h: 0.8 }, { fontSize: 17, valign: "middle" });
  X.disputed.slice(0, 3).forEach((d, i) => {
    const y0 = 2.3 + i * 1.5;
    s.addShape(pres.ShapeType.roundRect, { x: M, y: y0, w: W - 2 * M, h: 1.4, rectRadius: 0.08, fill: { color: C.background1 }, line: { color: "E1E0D9", width: 1 }, objectName: "dispute-card" });
    s.addText(String.fromCharCode(65 + i), { x: M + 0.15, y: y0 + 0.1, w: 0.5, h: 0.5, fontSize: 20, bold: true, color: C.accent2, margin: 0, isTextBox: true, objectName: "dispute-letter" });
    s.addText(`(${d.year}) ${d.title}`, { x: M + 0.7, y: y0 + 0.07, w: W - 2 * M - 0.9, h: 0.52, fontSize: 11, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "dispute-title" });
    s.addText([
      { text: "Human: ", options: { bold: true, color: H.dk1 } }, { text: d.gold + "        ", options: { color: H.dk1 } },
      { text: "Model: ", options: { bold: true, color: H.accent1 } }, { text: d.model, options: { color: H.accent1 } },
    ], { x: M + 0.7, y: y0 + 0.6, w: W - 2 * M - 0.9, h: 0.3, fontSize: 12, margin: 0, valign: "middle", isTextBox: true, objectName: "dispute-labels" });
    s.addText([
      { text: "Model's reason: ", options: { bold: true, color: H.accent6 } }, { text: (d.why.length > 230 ? d.why.slice(0, 227) + "..." : d.why), options: { color: H.accent6, italic: true } },
    ], { x: M + 0.7, y: y0 + 0.9, w: W - 2 * M - 0.9, h: 0.45, fontSize: 10.5, margin: 0, valign: "top", isTextBox: true, objectName: "dispute-reason" });
  });
  s.addNotes("A: a Defense Health Program appropriation. The humans say Health, the model follows the 'defense' word to Other. B: cigarette fire-safety standards; the humans say Other (consumer safety), the model hears tobacco. C: a Social Security marriage rule; the humans file family law under Law and Crime, the model under Social Welfare. Let them argue. Errors have three owners and each one gets a different fix: a prompt rule, a codebook clarification, or a correction of the gold.");

  // 18. Differential error
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Part 1: Labels as measurements" });
  s.addText("The error is not the same for everyone", { placeholder: "title" });
  await fig(s, "f05_differential_error.png", { x: M, y: 1.35, w: 8.7, h: 5.3 }, { alt: "Health share by party and era, humans vs model" });
  bullets(s, [
    { text: `Counted with the human codes: ${pct(BP.Democrat.health_gold)} of Democratic bills and ${pct(BP.Republican.health_gold)} of Republican bills are about health. No gap.`, bold: true },
    { text: `Counted with the model's labels, same bills: ${pct(BP.Democrat.health_llm)} vs ${pct(BP.Republican.health_llm)}. A ${pts(BP.Democrat.health_llm - BP.Republican.health_llm)}-point gap that does not exist.`, bold: true },
    `Accuracy is almost identical for both parties (${pct(BP.Democrat.accuracy, 0)} vs ${pct(BP.Republican.accuracy, 0)}). The direction of the errors differs, and that is what bends a comparison.`,
    `Over time the error changes sign: the model over-counts health in 1973-86 and in 2001-16 and under-counts it in between. The humans' series rises and levels off; the model's keeps climbing, and its accuracy slips from ${pct(BE["1973-1986"].accuracy, 0)} to ${pct(BE["2001-2016"].accuracy, 0)} across the eras.`,
  ], { x: 9.5, y: 1.5, w: 3.3, h: 5.0 }, { fontSize: 12, space: 8 });
  s.addNotes("This is the slide that answers 'is 88% good enough?'. Same accuracy by party, different sign of the error, and a partisan finding appears from nowhere. Always validate by the groups you will compare.");

  // 19. The estimation problem
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Part 1: Labels as measurements" });
  s.addText("Measure the error, then subtract it", { placeholder: "title" });
  card(s, { x: M, y: 1.45, w: 5.9, h: 2.4 }, "The scale analogy", "A scale reads heavy. You would not throw it away. You would weigh a few people on a good scale as well, measure how far off it is on average, and subtract that from the group's average weight. You correct the average, not each person.", C.accent2, { bodySize: 13.5 });
  card(s, { x: M + 6.2, y: 1.45, w: 5.9, h: 2.4 }, "The same move with labels", "Take a random subset of titles and have people code them. On that subset you see both the model's label and the truth, so you can measure the model's average error and subtract it from the model's average over all titles.", C.accent1, { bodySize: 13.5 });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 4.1, w: W - 2 * M, h: 1.25, rectRadius: 0.08, fill: { color: "1B1F23" }, line: { color: "1B1F23", width: 0 }, objectName: "formula-box" });
  s.addText([
    { text: "corrected share  =  ", options: { color: "F0EFEC", fontSize: 20 } },
    { text: "mean of model labels over all N titles", options: { color: "6DA7EC", fontSize: 20 } },
    { text: "  minus  ", options: { color: "F0EFEC", fontSize: 20 } },
    { text: "mean of (model label minus human label) over the n hand-coded titles", options: { color: "F9A27B", fontSize: 20 } },
  ], { x: M + 0.3, y: 4.15, w: W - 2 * M - 0.6, h: 1.15, valign: "middle", margin: 0, isTextBox: true, align: "center", objectName: "formula" });
  bullets(s, [
    "A label becomes a number: 1 if the title was called Health, 0 if not. The mean of those numbers is the share of health bills. The interval is the range the true share is likely to be in, given how much the n hand-coded titles can tell you.",
    "This is prediction-powered inference, PPI (Angelopoulos, Bates, Fannjiang, Jordan and Zrnic, Science 2023, from Berkeley), and the design-based supervised learning of Egami, Hinck, Stewart and Wei. One line of arithmetic; it holds whatever the model's error pattern, as long as the hand-coded subset is a random sample.",
  ], { x: M, y: 5.5, w: W - 2 * M, h: 1.4 }, { fontSize: 13, space: 6 });
  s.addNotes("The whole method in one line. Say the names only after the arithmetic. The condition that matters is the random subset: pick the hand-coded titles at random, never the easy ones.");

  // 20. Three estimators by party
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Part 1: Labels as measurements" });
  s.addText("Three estimates of one share: the finding dissolves", { placeholder: "title" });
  await fig(s, "f06_three_estimators_party.png", { x: M, y: 1.35, w: 8.8, h: 5.3 }, { alt: "Naive, human-only and corrected estimates by party" });
  bullets(s, [
    { text: `Naive: ${pct(E.Democrat.naive.est)} vs ${pct(E.Republican.naive.est)}, a ${pts(E.Democrat.naive.est - E.Republican.naive.est)}-point gap with intervals that ignore the model's error. Marginal on 1,650 titles; on 20,000 documents the same gap would look like a strong finding, far outside any naive interval, because naive intervals shrink with N and bias does not.`, bold: true },
    `Hand-coded subset only (${R.n_hand} titles): ${pct(E.Democrat.human.est)} vs ${pct(E.Republican.human.est)}. Right on average, but any one sample of ${R.n_hand} can land a point or two off, which is what the wide interval says.`,
    { text: `Corrected: ${pct(E.Democrat.corrected.est)} vs ${pct(E.Republican.corrected.est)}, with an interval about ${(E.Democrat.corrected.se / E.Democrat.human.se * 100).toFixed(0)}% as wide as the subset-only one. The gap is gone.`, bold: true },
    "The black mark is what all 1,650 hand-coded titles say. We usually cannot afford it; here it lets us check the method.",
  ], { x: 9.6, y: 1.5, w: 3.2, h: 5.0 }, { fontSize: 12.5, space: 9 });
  s.addNotes("Predict before you show the results: ask which way the correction will move the Democratic estimate. Then the three colors. The corrected interval is narrower than the human-only one because the model labels on the other 1,450 titles are not thrown away; they are used, with their error measured.");

  // 21. Over time
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Part 1: Labels as measurements" });
  s.addText("Back to the opening chart: which part was the instrument?", { placeholder: "title" });
  await fig(s, "f07_three_estimators_era.png", { x: M, y: 1.35, w: 8.8, h: 5.3 }, { alt: "Naive, human-only and corrected estimates by era" });
  bullets(s, [
    { text: `Humans: ${pct(E["1973-1986"].truth)}, ${pct(E["1987-2000"].truth)}, ${pct(E["2001-2016"].truth)}. Model: ${pct(E["1973-1986"].naive.est)}, ${pct(E["1987-2000"].naive.est)}, ${pct(E["2001-2016"].naive.est)}. After 2000 the two series part ways on the same bills: the model over-counts by 2 points where it under-counted before. That divergence is well beyond noise. The climb after 2000 is the instrument, not the bills.`, bold: true },
    `The correction is honest about what ${E["1973-1986"].n_hand} to ${E["1987-2000"].n_hand} hand-coded titles per era can tell you: intervals of plus or minus ${pts(eraHWmin, 0)} to ${pts(eraHW, 0)} points that cover the truth in every era but cannot tell a plateau from a climb.`,
    `In 1973-86 the corrected point even lands farther from the truth than the naive one, and the subset alone found ${Math.round(E["1973-1986"].human.est * E["1973-1986"].n_hand)} health bill in ${E["1973-1986"].n_hand}: an interval from ${pct(E["1973-1986"].human.ci_low)} to ${pct(E["1973-1986"].human.ci_high)}, honest and useless. An interval that admits its width is the answer; a point alone would mislead. That is the price of splitting 200 titles three ways. How many would you need? Challenge 3.`,
  ], { x: 9.6, y: 1.5, w: 3.2, h: 5.0 }, { fontSize: 12, space: 8 });
  s.addNotes("Close the loop with the hook: the rise is real, the steady climb after 2000 is the instrument drifting. Then be honest about the correction before anyone else is: with 65 titles per era the intervals are wide and the 1973-86 point overshoots. An interval that admits it is worth more than a point that pretends. That is the best argument for Challenge 3.");

  // 22. Challenge 3
  s = pres.addSlide({ masterName: "TINT", sectionTitle: "Part 1: Labels as measurements" });
  s.addText("Challenge 3: how many titles would you code by hand?", { placeholder: "title" });
  numCircle(s, 3, M, 1.4, H.accent1);
  para(s, "Your project: the same instrument on 20,000 bill titles. Your committee wants to know whether Democrats and Republicans differ by 3 points or more in the share of health bills. A research assistant codes about 100 titles an hour. Rule: to tell a 3-point gap between two groups from zero, each group's interval must be about plus or minus 2 points or narrower. Read the curve: how many titles per party, how many RA hours, and what interval do you promise?", { x: M + 0.7, y: 1.3, w: 11.5, h: 1.15 }, { fontSize: 13.5, valign: "middle" });
  await fig(s, "f08_sample_size_curve.png", { x: M, y: 2.55, w: 7.6, h: 4.2 }, { alt: "Interval width against number of hand-coded titles, with the target line" });
  bullets(s, [
    { text: `Corrected: about ${nC3} titles per party reaches plus or minus 2, so ${2 * nC3} titles and ${Math.round(2 * nC3 / 100)} RA hours. Hand-coding alone would need about ${nH3} per party, ${Math.round(2 * nH3 / 100)} hours.`, bold: true },
    `With 200 per party the interval is plus or minus ${pts(PJrow[200].ci_halfwidth)}, so a 3-point gap could not be told from zero. With 100, plus or minus ${pts(PJrow[100].ci_halfwidth)}.`,
    "Doubling n cuts the width by about 30%; four times the work buys half. The first 200 titles are worth more than the next 600.",
    "The curve comes from this lesson's error rates. In your project, build it from a pilot of 100 hand-coded documents, then decide.",
    "Floor: 200 at random for one share. A comparison between groups costs more per group (here, about 400). A corpus smaller than 200 documents (40 interviews)? Hand-code it all; the method is for collections too large to read.",
  ], { x: 8.5, y: 2.6, w: 4.3, h: 4.2 }, { fontSize: 11.5, space: 6 });
  s.addNotes("The exit ticket: each person writes n per party, RA hours and the interval, on the handout. The checkable answer is about 400 per party, 800 titles, one day of RA time. 200 per party, the lesson's floor, is the wrong answer to this question, and that is the point: the floor is for one share, a comparison costs more. Four minutes, including the reveal.");

  // ================================================================ PART 2: LINKAGE
  pres.addSection({ title: "Part 2: Links as measurements" });

  // 23. Same recipe
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Part 2: Links as measurements" });
  s.addText("Second problem, same recipe: is this the same person?", { placeholder: "title" });
  await fig(s, "f09_linkage_score_band.png", { x: M, y: 1.35, w: 7.6, h: 4.6 }, { alt: "Candidate pairs by number of agreeing fields" });
  bullets(s, [
    "Febrl: two files of 5,000 person records; file B is a corrupted copy of file A (typos, swapped digits, moved addresses), so the 5,000 true links are known.",
    "Four fields two administrative files would share: given name, surname, date of birth, postcode. No ID. Your version: census manuscripts across decades, or hospital discharges to death records.",
    `Comparing every record with every other would mean 25 million pairs. Blocking compares only pairs that share something (a similar surname, the same birth date or postcode): ${BL.candidates.toLocaleString()} candidate pairs, and only ${BL.lost_links} true links lost. All ${BL.lost_links} are people who moved: the lesson's own mechanism, already at work in the blocking step.`,
    `Rules sort most pairs: ${BL.sure.toLocaleString()} sure matches, ${(BL.candidates - BL.sure - BL.band).toLocaleString()} sure non-matches, and a band of ${BL.band.toLocaleString()} ambiguous pairs with ${BL.band_true} true links inside. At scale the rules become Fellegi-Sunter probabilistic weights, the band is their uncertain middle, and deduplication within one file is the same problem.`,
  ], { x: 8.5, y: 1.5, w: 4.3, h: 5.0 }, { fontSize: 12, space: 7 });
  s.addNotes("Linking records is classification too: the unit is the pair, the label is same person or not. Rules do the cheap part; the band is where a human or a model has to decide. Keep this to two minutes.");

  // 24. Policies
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Part 2: Links as measurements" });
  s.addText("Four ways to handle the band, measured against the truth", { placeholder: "title" });
  await fig(s, "f10_linkage_policies.png", { x: M, y: 1.35, w: 8.0, h: 4.6 }, { alt: "Precision and recall of four linkage policies" });
  bullets(s, [
    `Drop the band: precision ${pct(LP["Rules, strict (drop the band)"].precision)}, recall ${pct(LP["Rules, strict (drop the band)"].recall)}. Clean, but 14% of the people are gone.`,
    `Keep the band: recall ${pct(LP["Rules, lenient (keep the band)"].recall)}, precision ${pct(LP["Rules, lenient (keep the band)"].precision)}. One in six links is invented.`,
    `Let the model adjudicate the band: precision ${pct(LP["Rules + Haiku on the band"].precision)}, recall ${pct(LP["Rules + Haiku on the band"].recall)}. Inside the band it finds ${pct(R["band_claude-haiku-5-5"].recall_in_band, 0)} of the true pairs and is right ${pct(R["band_claude-haiku-5-5"].precision_in_band, 0)} of the time it says yes.`,
    "The model costs cents and replaces a day of a person checking pairs by hand (clerical review). It is still a measurement: we only know these numbers because the truth is known here. In your project, you hand-review a random sample of the band (that gives precision) and of the near-misses just below it, where the missed links hide (that gives recall): the human-in-the-loop review any automated pipeline needs before anyone trusts it.",
  ], { x: 8.9, y: 1.5, w: 3.9, h: 5.0 }, { fontSize: 12.5, space: 8 });
  s.addNotes("Same vocabulary as part one: precision and recall, a validation sample, a model that is good but not free of error. The stronger model is not better here; say it.");

  // 25. Downstream + question
  s = pres.addSlide({ masterName: "TINT", sectionTitle: "Part 2: Links as measurements" });
  s.addText("Question: would you rather lose pairs or invent them?", { placeholder: "title" });
  await fig(s, "f11_linkage_downstream.png", { x: M, y: 1.35, w: 7.6, h: 4.3 }, { alt: "Share who changed postcode under each linkage policy" });
  bullets(s, [
    { text: `The research question: what share of people changed postcode? Truth: ${pct(MV["Truth (all true links)"])}.`, bold: true },
    `Drop the band: ${pct(MV["Rules, strict"])}. Keep it: ${pct(MV["Rules, lenient"])}. Model adjudication: ${pct(MV["Rules + Haiku"])}.`,
    "Postcode is both a matching field and the outcome, so the linkage error is correlated with what you measure. Missed links hide movers; invented links create them.",
    "The answer depends on the number you want to know. For a mobility rate, missed matches push it down and false matches push it up. Febrl is synthetic (the moves were planted, which is why every true link is known); in your data the truth is unknown, so you hand-review random samples of the band and of the near-misses and measure both error rates, exactly like part one.",
  ], { x: 8.5, y: 1.5, w: 4.3, h: 5.0 }, { fontSize: 12.5, space: 9 });
  s.addNotes("Let them answer. There is no free policy: strict loses the movers, lenient manufactures them. Then the link to part one: the band is your validation sample, review a random piece of it, measure both error rates by group, correct.");

  // ================================================================ CLOSE
  pres.addSection({ title: "Close" });

  // 26. What we know works
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Close" });
  s.addText("What we know works (as of 2026)", { placeholder: "title" });
  const works = [
    ["Codebook in the prompt", "Definitions and boundary rules. In this lesson they were worth 14 points of accuracy; the examples, nothing measurable; the model upgrade, one and a half.", C.accent1],
    ["Structured output, fixed settings", "A fixed JSON answer, a one-sentence reason you can read, and the same settings for every call, so a label today and a label next month come from the same instrument.", C.accent1],
    ["Log the instrument", "Model name and version, prompt version, date, tokens. A changed prompt is a new instrument.", C.accent1],
    ["Random validation sample", "Hand-code a random subset; report accuracy, kappa, precision and recall by class and by the groups you compare.", C.accent2],
    ["Measure stability", "The same title can get a different label on a second run. Run a random sample twice and report run-to-run agreement next to accuracy.", C.accent2],
    ["Correct the estimate", "Prediction-powered inference or its design-based relatives; never treat AI labels as the truth in a regression or a trend.", C.accent3],
  ];
  works.forEach((w, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    card(s, { x: M + col * 4.1, y: 1.45 + row * 2.45, w: 3.95, h: 2.3 }, w[0], w[1], w[2], { bodySize: 12.5, titleSize: 14.5 });
  });
  s.addNotes("This is the list a student can take to a methods appendix. Six items, three about the instrument, two about validation, one about estimation. Nothing here depends on which model exists next year.");

  // 27. Report, privacy, cost
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Close" });
  s.addText("Before you send data to a model: report, privacy, cost", { placeholder: "title" });
  card(s, { x: M, y: 1.45, w: 3.95, h: 4.9 }, "What to report", "Codebook and prompt (verbatim, in an appendix). Model and version, date, settings. Size and selection of the validation sample. Accuracy, kappa, precision and recall by class and by group. How the estimate was corrected, and the interval. Where the cached outputs and code live.", C.accent1, { bodySize: 13 });
  card(s, { x: M + 4.1, y: 1.45, w: 3.95, h: 4.9 }, "Privacy", "Text sent to a model over the internet (an API call) leaves your machine. Interview transcripts, student records and anything under IRB need a data use review first; a model that runs on your own machine, a campus-hosted service, or de-identification before the call are the usual routes. With restricted data (an FSRDC project, a data use agreement) nothing leaves the enclave: you validate inside it, and disclosure review decides which cells of your error table come out. Bill titles are public; your data may not be.", C.accent2, { bodySize: 13 });
  card(s, { x: M + 8.2, y: 1.45, w: 3.95, h: 4.9 }, "Cost and time", `This lesson's full run: 6,600 title labels and 2,966 pair decisions, ${((Object.values(R.tokens_titles).reduce((a, b) => a + b.in, 0) + Object.values(R.tokens_pairs).reduce((a, b) => a + b.in, 0)) / 1e6).toFixed(1)} million input tokens, under 20 dollars and about an hour. The hand-coded sample of 200 titles: an afternoon. The afternoon is what makes the dollars usable.`, C.accent3, { bodySize: 13 });
  s.addNotes("Thirty seconds each. The privacy card is where the D-Lab's consulting questions land: 'can I put my interviews into ChatGPT' is a question about this card.");

  // 28. Key points + take-home
  s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Close" });
  s.addText("Key points", { placeholder: "title" });
  const kps = [
    "An AI label or link is a measurement with error. Treat the prompt as the instrument and freeze it before scoring.",
    "Validate on a random hand-coded sample, by class and by the groups you will compare. Accuracy alone can hide a spurious finding.",
    "High accuracy does not make an estimate unbiased. Measure the error and subtract it; prediction-powered inference does this in one line.",
    "How many documents to hand-code is a design decision read off a curve: 200 at random is the floor for one share, and a comparison between groups costs more per group.",
    "The same recipe handles record linkage: rules for the easy pairs, a model for the band, a reviewed sample for the error. The downstream number moves with the policy, so report the policy with the number.",
  ];
  kps.forEach((t, i) => {
    numCircle(s, i + 1, M, 1.5 + i * 0.82, i < 4 ? H.accent1 : H.accent3);
    para(s, t, { x: M + 0.7, y: 1.5 + i * 0.82, w: 11.4, h: 0.75 }, { fontSize: 15.5, valign: "middle" });
  });
  s.addShape(pres.ShapeType.roundRect, { x: M, y: 5.72, w: W - 2 * M, h: 1.1, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: "E1E0D9", width: 0.75 }, objectName: "takehome-box" });
  s.addText([
    { text: "Take-home challenge: ", options: { bold: true, color: H.accent2 } },
    { text: "take 1,000 documents from your own project, hand-code 200 chosen at random, run the notebook's three estimates of a share you care about, and bring the corrected interval to a D-Lab consultation. Fewer than 200 documents? Code them all and skip the model.", options: { color: H.dk1, breakLine: true } },
    { text: "Slides, notebooks, data and the full analysis: ", options: { bold: true, color: H.accent1 } },
    { text: "github.com/felipehrv/validate-before-you-estimate", options: { color: H.accent1, hyperlink: { url: "https://github.com/felipehrv/validate-before-you-estimate" } } },
  ], { x: M + 0.25, y: 5.78, w: W - 2 * M - 0.5, h: 1.0, fontSize: 13.5, valign: "middle", margin: 0, isTextBox: true, objectName: "takehome" });
  s.addNotes("Read the five points slowly; they are the exit ticket. The take-home challenge connects the workshop to D-Lab consulting, which is where the lesson pays off.");

  // 29. Thank you
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Close" });
  s.addText("Thank you", { placeholder: "title" });
  s.addText("Questions about the lesson", { placeholder: "body" });
  para(s, "Materials: notebooks (lessons and solutions), data, prompts, cached model outputs and this deck are available on request.\nFelipe H. Rodriguez Villalta  |  felipe.hrv@gmail.com", { x: M, y: 5.6, w: W - 2 * M, h: 1.0 }, { fontSize: 14, color: "CADCFC" });
  para(s, "Sources: Comparative Agendas Project (comparativeagendas.net); Febrl datasets via the recordlinkage package; Angelopoulos et al., Science 2023; Egami, Hinck, Stewart and Wei, NeurIPS 2023 and 2024; Desai, Card and Jacobs, arXiv 2026 (validation practice in eight social science journals); D-Lab workshop template and jupyter style guide.", { x: M, y: 6.6, w: W - 2 * M, h: 0.6 }, { fontSize: 10.5, color: "898781" });
  s.addNotes("Stop talking. Smile. Wait for questions.");

  const out = path.join(ROOT, "slides", "Validate_Before_You_Estimate_DLab_Lesson.pptx");
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("written", out);
})();
