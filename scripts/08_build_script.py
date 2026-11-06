"""
Build the presenter's guide and spoken script for the mock lesson:
  docs/Presentation_Script.docx
For every slide: a short note in Spanish on what the slide is about and what
the point is (for the presenter), then the spoken text in English, with the
clock time at which the slide should start. Numbers come from
outputs/results.json so the script cannot drift from the deck.
Run: python3 scripts/08_build_script.py
"""

import json
import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

PROJECT = Path(__file__).resolve().parents[1]
R = json.load(open(PROJECT / "outputs/results.json"))
E = R["estimates"]
S = {f"{r['model']}|{r['prompt']}": r for r in R["showdown"]}
BP = R["by_party_haiku_p3"]
BE = R["by_era_haiku_p3"]
LP = R["linkage_policies"]
MV = R["moved_rate"]
BL = R["blocking"]
PJ = R["sample_size_curve_projected"]
PJrow = {r["n_hand_coded"]: r for r in PJ["rows"]}
CM = R["confusion_haiku_p3"]          # keyed [model label][human label]
PC = {r["label"]: r for r in R["per_class_haiku_p3"]}
h1, h2, h3, s3 = (S["claude-haiku-5-5|prompt_1_names"], S["claude-haiku-5-5|prompt_2_codebook"],
                  S["claude-haiku-5-5|prompt_3_codebook_examples"], S["claude-sonnet-5-5|prompt_3_codebook_examples"])
N_C3 = -(-PJ["n_for_target_corrected"] // 50) * 50
N_H3 = -(-PJ["n_for_target_human_only"] // 50) * 50

p1 = lambda v: f"{100 * v:.1f}"
p0 = lambda v: f"{100 * v:.0f}"
INK = RGBColor(0x1B, 0x1F, 0x23)
MUTED = RGBColor(0x52, 0x51, 0x4E)
BLUE = RGBColor(0x2A, 0x78, 0xD6)
GUIDE_FILL = "EEF4FC"

# ---------------------------------------------------------------------------
# For each slide: (number, title, clock, guide in Spanish, script in English).
# [Brackets] in the script are stage directions, not to be read.
# ---------------------------------------------------------------------------
SCRIPT = [
("Part A. Overview (0:00 to 8:00)", [
(1, "Title", "0:00",
"""De qué va este slide: tu presentación personal y el plan de los 40 minutos. El panel necesita dos cosas desde el inicio: quién sos en una línea, y que primero viene un overview de 8 minutos y después enseñás con ellos como estudiantes. Repartís el handout acá para no cortar el ritmo más adelante.""",
"""
Good afternoon, and thank you for having me. I am Felipe Rodriguez Villalta, a research data analyst at Stanford. For about ten years my work has been to make research data usable, at UCSF, the ILO, the World Bank and now Stanford, and I teach the way I work.

Here is the plan for our time together. First, about eight minutes of overview: what the lesson is for, the learning objectives, the exercises and the teaching approach. Then I will teach most of the lesson with you as my graduate students, for about thirty minutes. After that I am happy to take your questions.

[Hand out the sheet.] This is the handout for the exercises. Please keep it in front of you; we will use it three times.
"""),
(2, "Why this lesson, and why now", "0:45",
"""De qué va: por qué existe la lección. Tres hechos: la gente ya usa modelos de lenguaje para codificar texto; una revisión reciente (Desai, Card y Jacobs, 2026) muestra que casi nadie valida bien esas etiquetas; y Egami y coautores demostraron que si tratás las etiquetas del modelo como verdad, tus estimaciones salen sesgadas y tus intervalos de confianza dejan de valer, aunque el modelo acierte más del 90%. De ahí la única idea de la clase: una etiqueta de IA es una medición con error; hay que medir ese error contra personas y llevarlo a la estimación.""",
"""
Why does this lesson exist? Three reasons.

First, the practice is already here. Graduate students use large language models to code interviews, open-ended survey answers, documents and social media posts. It is cheap, it is fast, and often it is good.

Second, the checking is not keeping up. This year, Desai, Card and Jacobs reviewed twenty-seven papers in eight top social science journals. In those papers, the labels produced by language models were central to the analysis, and the way the authors checked those labels against human judgment was limited and inconsistent.

Third, this is not a cosmetic problem. Egami and coauthors have shown that if you treat the model's labels as if they were true, your estimates come out biased and your confidence intervals stop meaning what they claim, even when the model is right more than ninety percent of the time.

So this lesson teaches one idea, and it is on the screen. Every label an AI gives you is a measurement with error. Measure that error against people, and then carry it into your estimate.
"""),
(3, "Learning objectives", "2:00",
"""De qué va: los cinco objetivos de aprendizaje, que el D-Lab pidió explícitamente. Cada uno se practica en un ejercicio. El tercero es el que distingue esta clase de cualquier otra sobre LLMs: una exactitud alta no garantiza una estimación sin sesgo.""",
"""
These are the five learning objectives. Each one is practiced in an exercise later, so I will read them quickly.

One: explain why a classification prompt is a measurement instrument, and what changes when you put the codebook inside it. Two: design a validation sample and read the agreement numbers as measurement error. Three: explain why high accuracy does not guarantee an unbiased estimate, or a fair comparison between groups. Four: read a corrected estimate, and choose how many documents to code by hand using a curve. Five: apply the same recipe to record linkage.

Objective three is the heart of the lesson. Most AI workshops stop at objective two: they teach you to measure accuracy. This lesson asks what that accuracy does to your results.
"""),
(4, "The lesson at a glance", "3:00",
"""De qué va: la línea de tiempo de la lección completa de 45 minutos y la tabla con los cuatro ejercicios y la meta de cada uno (lo que el D-Lab llama "exercise goals"). Decí los bloques y una frase por ejercicio; la tabla está en pantalla, no hace falta leerla entera.""",
"""
Here is the lesson at a glance: forty-five minutes, six blocks, four exercises. Five minutes for an opening chart and a poll. Eight minutes to go from a codebook to a prompt. Eleven minutes to validate the labels. Eight minutes to correct the estimate. Eight minutes for record linkage. And five minutes to close.

The four exercises are in the table, and each one has a goal stated as something the participant can do afterward. In Challenge one, participants code six real bill titles by hand and compare with a neighbor. The goal is to feel how ambiguous a codebook is, and to see that even people do not agree one hundred percent of the time. In Challenge two, they bet on which of three prompts works best, and then they vote on three titles where the model and the humans disagree. The goal is to separate the prompt from the model as sources of error, and to decide whose error it is. In Challenge three, they read a curve and choose how many titles to hand-code per group, and that one has an answer I can check. And in part two the question is: would you rather lose record pairs or invent them? The goal is to see that a linkage policy is a decision about your estimate, not a technical detail.

That is an exercise every ten to twelve minutes. The D-Lab guide asks for one every fifteen.
"""),
(5, "Teaching approach", "4:15",
"""De qué va: las seis decisiones de diseño de la clase. Son los valores pedagógicos del propio D-Lab (hands-on, sin código, ética y evaluación dentro del método, evaluación formativa), dichos como decisiones tuyas y con su razón. Una frase por tarjeta.""",
"""
Six design choices, with the reason for each.

Backward design: I started from the decisions a student has to make in a real project, how many documents to hand-code and what to report, and I built the exercises around those decisions.

Predict, observe, explain: before I show any result, participants commit to a guess. The distance between the guess and the data is what they remember.

Hands-on without a key: every call to the model was run beforehand and saved, so the notebook runs on DataHub or Binder without an API key, or on a laptop with no internet at all.

A path without code: hand-coding, voting and reading a curve need no Python, so a student from history or anthropology can take the whole lesson.

Evaluation and ethics inside the method: validation, error by group, privacy and cost are steps of the workflow, not a slide at the end.

And formative assessment: a poll, quick check questions during the lesson, and Challenge three as an exit ticket with a number I can check.
"""),
(6, "Where it sits in the D-Lab catalog", "5:45",
"""De qué va: dónde se ubica la lección en el catálogo del D-Lab. Dos talleres de ellos que ya existen (Python APIs for LLMs enseña la mecánica; Qualitative Coding with AI Tools enseña a auditar los códigos de un agente), esta lección en el medio como el paso de medición, y el taller siguiente (record linkage completo). Después, qué es nuevo: el método viene de Berkeley y Princeton; lo tuyo es la versión de 45 minutos sin código, la decisión de presupuesto y la aplicación al linkage. (Las frases de prerrequisitos y "not covered" se sacaron del slide y del guion, según tu marca en rojo.)""",
"""
Where does this sit in the D-Lab catalog? Python APIs for Large Language Models teaches the mechanics: how to call a model and get a structured answer back. Qualitative Coding with AI Tools teaches the audit: you code a few passages by hand, you run the AI twice, and you check its codes against the text. This lesson is the measurement step between them: how large the error is, and what it does to your estimate. The natural next workshop is record linkage in full.

What is new here? The method itself comes from Berkeley, prediction-powered inference, and from Princeton, design-based supervised learning. This lesson turns it into a forty-five minute class with no code, with a practical budget decision, and applies the same recipe to record linkage.
"""),
(7, "Materials and what participants need", "6:45",
"""De qué va: los materiales. Los datos (títulos de proyectos de ley con códigos humanos; registros de personas con los links verdaderos conocidos), el workflow de IA ya corrido y guardado con su registro, el repositorio en el template del D-Lab que un Fellow puede dictar, y el handout. Treinta segundos.""",
"""
The materials, in thirty seconds. The data: one thousand six hundred and fifty bill titles from the Comparative Agendas Project, seventy-five per Congress from 1973 to 2016, each one coded by trained human coders. And a second dataset, Febrl: two files of five thousand person records, where we know which records belong to the same person.

The AI workflow: three prompts, two models, six thousand six hundred labels and three thousand decisions on record pairs, all run beforehand and saved with the model name, the prompt version, the tokens and the time. The whole run cost under twenty dollars.

The repository follows the D-Lab template, with lesson and solution notebooks, and runs on DataHub and Binder, so a Data Science Fellow could teach it after one rehearsal. In the room we only need this handout and the slides.
"""),
(8, "Transition", "7:30",
"""De qué va: la bisagra. Acá dejás de hablar como candidato y pasás a ser el instructor, con el panel como estudiantes. Tres preguntas que van a poder responder al final, y la regla del D-Lab: está bien no saber.""",
"""
That is the overview. Now I change hats. [Pause.]

For the next thirty minutes, you are my graduate students. When we are done, you will be able to answer three questions. Can I trust this label? Can I trust this estimate? And how many documents do I have to code by hand?

One rule for the room, and it is the D-Lab's own: it is OK not to know. Every bet in this lesson is a guess, and the wrong guesses are the ones that teach.
"""),
]),
("Part B. Lesson, part one: labels as measurements (8:00 to 32:00)", [
(9, "Hook", "8:00",
f"""De qué va: el gancho. Un gráfico que parece un hallazgo: el porcentaje de proyectos de ley sobre salud sube de {p1(BE['1973-1986']['health_llm'])}% a {p1(BE['2001-2016']['health_llm'])}% en tres períodos. El gráfico es real pero las etiquetas las puso un modelo con 88% de exactitud. No digas todavía qué tiene de malo: la respuesta llega en el slide 22 (la subida después del 2000 es el instrumento, no los proyectos de ley). La encuesta a mano alzada muestra que casi nadie valida: muchas manos en la primera pregunta, casi ninguna en la tercera.""",
f"""
Let me start with a chart. Would you put this chart in your dissertation?

It shows the share of all bills in Congress that are about health, in three periods of time. {p1(BE['1973-1986']['health_llm'])} percent in the first period, {p1(BE['1987-2000']['health_llm'])} in the second, {p1(BE['2001-2016']['health_llm'])} in the third. Health kept climbing on the agenda, period after period. It looks clean, and it looks like a finding.

Before I tell you what is behind it, a quick poll, hands up. Have you ever used an AI model to code text for research? [Pause.] Did you check its labels against a person? [Pause.] And did that check change a number in your results? [Pause.]

Usually many hands go up for the first question and very few for the third. That gap is what this lesson is about.

The chart is real. The labels came from a language model that is right eighty-eight percent of the time. By the end of the lesson you will know how much of that climb is real, and how much of it is the instrument.
"""),
(10, "The data: bill titles with a human answer key", "9:30",
"""De qué va: los datos. El Comparative Agendas Project (CAP) codificó todos los proyectos de ley del Congreso desde 1947 en 21 temas, con codificadores humanos entrenados; eso nos da la "clave de respuestas". Tomamos 1,650 títulos (75 por Congreso, 1973 a 2016) y reducimos a 7 categorías. Dos advertencias que se usan después: "Other" es el 72% de los títulos (por eso la exactitud sola engaña: un modelo que diga siempre Other acierta 72% sin leer nada) y los humanos tampoco son perfectos (dos codificadores coinciden entre 90 y 95% de las veces). Cerrá con "your version": ejemplos de otras disciplinas para que cada uno vea su propio problema.""",
"""
First, the data. The Comparative Agendas Project, CAP, is an academic project that took every bill introduced in the U.S. Congress since 1947 and assigned each one a policy topic, using trained human coders. That gives us something rare: an answer key. For every bill title, we know what the humans said.

We keep seven topics: Health, Education, Environment, Immigration, Law and Crime, Social Welfare, and Other. We take seventy-five titles from each Congress between 1973 and 2016, one thousand six hundred and fifty titles in total. Each title also comes with the party of the legislator who sponsored it.

Two things to keep in mind, because we will use them later. First, "Other" is seventy-two percent of the titles. That means a model that simply answered "Other" for every title, without reading anything, would already be right seventy-two percent of the time. So a high accuracy number, by itself, does not tell you much. Second, the human codes are our answer key, but they are not perfect either. When two trained coders code the same bill, they agree ninety to ninety-five percent of the time. So even our gold standard has some error in it.

And your version of this data could be: interview transcripts coded for housing insecurity, by gender of the respondent. Nineteenth-century newspaper ads coded for occupation, by decade. Clinic notes coded for tobacco use, by site. The structure is the same.
"""),
(11, "From codebook to prompt", "10:30",
"""De qué va: los tres prompts como tres "instrumentos de medición". P1 le da al modelo solo los nombres de las categorías; P2 agrega las definiciones del codebook (con las reglas para los casos difíciles); P3 agrega además 14 ejemplos codificados por humanos, tomados fuera de la muestra que se evalúa (así el modelo nunca ve un ejemplo que después se califica). Todas las llamadas devuelven el mismo formato (una etiqueta y una frase de razón) y quedan registradas. "Frozen" quiere decir: el prompt se escribe, se cierra, y recién después se corre y se compara; no se va retocando mirando los resultados. "Instrumento" quiere decir: como un cuestionario de encuesta, si cambiás la pregunta cambiás lo que medís; por eso cada versión del prompt necesita su propia validación.""",
"""
From that codebook I wrote three prompts, three sets of instructions for the model. I want you to think of each one as a version of a measurement instrument, the way a survey questionnaire is an instrument.

P1 gives the model only the names of the seven topics: classify this title as Health, Education, and so on. Nothing else. The model has to fill in the definitions from whatever it already knows.

P2 adds a condensed version of the CAP codebook: for each topic, what belongs in it, including the rules for the difficult cases. For example, treatment for drug addiction is Health, but drug trafficking is Law and Crime, and programs for veterans are Other.

P3 adds, on top of that, fourteen example titles already coded by humans, two for each topic. Those fourteen were taken from outside the set of titles we evaluate, so the model never sees an example that is later scored.

Every call, for every prompt and every title, asks for the same fixed answer: one label, and one short sentence explaining why. Each answer is saved together with the model name, the prompt version, the number of tokens and the time. And one discipline matters here: the prompt is written, then locked, and only then run and compared with the human codes. We do not keep adjusting the prompt while looking at the scores, because then we would be tuning it to the test.

Why the word instrument? Because a prompt is a questionnaire for the model. If you change the wording of a survey question, you change what you measure. In the same way, a changed prompt is a changed instrument, and it needs its own validation.
"""),
(12, "Challenge 1: be the coder", "12:00",
"""De qué va: el primer ejercicio. Ellos codifican seis títulos reales con el codebook de siete temas que está en el handout, tres minutos, y comparan con el vecino. El punto es que sientan en carne propia que el codebook es ambiguo.""",
"""
Challenge one. Now you are the coder.

On your handout you have six real bill titles and the codebook with the seven topics. Code each title into exactly one topic. You have three minutes. Then compare with your neighbor, and circle the titles where the two of you disagree.

[Three real minutes. Walk around. Then ask:] Where did you and your neighbor disagree, and what rule would have settled it?
"""),
(13, "What the trained coders said", "15:00",
"""De qué va: los resultados del Challenge 1 (qué código les puso el CAP a los seis títulos) y tres lecciones. Primera: los títulos en los que dos personas no coinciden son los ambiguos, y son los mismos en los que el modelo va a diferir de los humanos; como las personas no coinciden al 100%, el modelo tampoco puede: el acuerdo entre humanos es el techo. Segunda: los títulos 3 y 4 son cambios a la ley de impuestos (uno exime del impuesto a la nómina a los trabajadores mayores de 65; el otro crea bonos para construir escuelas a través del código tributario). Como "impuestos" aparece bajo Other en el codebook, mucha gente los pone en Other, pero el CAP los codificó Social Welfare y Education, porque la regla es codificar para quién es el proyecto, no la herramienta que usa. Tercera: la clave de respuestas también tiene errores; el ejemplo es real: en los datos del CAP, la Torture Victims Relief Act (financia centros de tratamiento para víctimas de tortura) está codificada como Environment.""",
"""
Here is what the trained coders said. [Give them a moment to compare.]

The take-back program for controlled substances: Health. Retirement benefits for police officers and firefighters: Law and Crime. The payroll tax exemption for workers over sixty-five: Social Welfare. The bonds for school buildings: Education. Border security: Immigration. The national heritage areas: Other.

Three lessons from this exercise.

First, look at where you and your neighbor disagreed. Those are the ambiguous titles, and they are exactly the titles where the model will disagree with the humans too. And since two trained people do not agree one hundred percent of the time, we cannot expect the model to agree one hundred percent with the human codes either. The agreement between humans is the ceiling.

Second, titles three and four. Both of them are changes to the tax law: one exempts workers over sixty-five from the payroll tax, the other creates bonds for school construction through the tax code. In the codebook, taxes appear under Other, so many people put these two under Other. But the CAP coders put them under Social Welfare and Education. The reason is a rule the codebook states in one sentence: code the bill by who or what it is for, the older workers, the schools, not by the tool it uses, which here is a tax change. If you missed that rule, so did the model, and in a minute you will see that many of its errors are exactly this kind of slip into Other.

Third, our answer key is a measurement too. This is a real example from the CAP data: a bill that funds treatment centers for victims of torture is coded Environment. That is a human slip. Keep it in mind for Challenge two: when the model and the human disagree, the human is not always the one who is right.
"""),
(14, "Challenge 2, part one: place your bet", "16:30",
"""De qué va: Challenge 2, parte uno: apuestas antes de ver los resultados. Tres apuestas: cuál de los tres prompts acierta más y cuánto; si el modelo caro (Sonnet) supera al barato (Haiku) por más o por menos de 5 puntos; y cuánto acierta un "codificador vago" que siempre contesta Other. La tercera apuesta prepara la explicación de kappa en el slide siguiente.""",
"""
Challenge two, part one. Place your bet before you see any result.

Same one thousand six hundred and fifty titles, same cheap model, three prompts. Which of the three prompts agrees most with the human coders, and by how much? Mark your guess on the handout. [Show of hands for each prompt.]

Second bet. Now we switch to a stronger and more expensive model, Sonnet instead of Haiku, with the best prompt. Does it beat the cheap model by more than five points of accuracy, or by less?

Third bet, for the brave. Imagine a lazy coder who does not read anything and answers "Other" every single time. What accuracy would that coder get?
"""),
(15, "The results: the prompt moves the needle, the model barely does", "17:30",
f"""De qué va: los resultados de las apuestas. P1 {p1(h1['accuracy'])}%, P2 {p1(h2['accuracy'])}%, P3 {p1(h3['accuracy'])}%; Sonnet {p1(s3['accuracy'])}%. Conclusión: el codebook dentro del prompt vale 14 puntos; los ejemplos no agregan nada medible; el modelo caro, 1.5 puntos. Y el codificador vago: 72%, porque el 72% de los títulos son Other. Eso quiere decir que P1 (73%) apenas supera a no leer nada, y por eso se reporta kappa al lado de la exactitud: kappa mide cuánto mejor que el azar es el instrumento, de 0 a 1 (0 para el vago, {h3['kappa']:.2f} para P3). Mensaje: el prompt mueve la aguja, el modelo casi no; el instrumento es lo que controlás y lo que tenés que reportar.""",
f"""
Here are the results.

With the names only, P1, the model agrees with the human coders {p1(h1['accuracy'])} percent of the time. With the codebook, P2, {p1(h2['accuracy'])}. With the codebook plus examples, P3, {p1(h3['accuracy'])}. So putting the definitions into the prompt was worth fourteen points of accuracy. Adding the examples added almost nothing, a third of a point, which is within the noise you get from running the same model twice. And the stronger model, Sonnet, adds one and a half points over the cheap one, and that difference is real.

Now the lazy coder, the third bet. A coder who answers "Other" every single time is right seventy-two percent of the time, simply because seventy-two percent of the titles are Other. Notice what that means: our first prompt, at seventy-three percent, was barely better than not reading the titles at all. That is why we report a second number next to accuracy, kappa. Kappa measures how much better than chance an instrument is, on a scale from zero to one. The lazy coder gets a kappa of zero. P3 gets {h3['kappa']:.2f}.

The message of this slide: the prompt moves the needle, and the model barely does. The instrument is the thing you control, and it is the thing you must report.
"""),
(16, "Where the disagreements live", "19:00",
"""De qué va: la matriz de confusión, es decir, en qué categorías se equivoca el modelo. Cada fila es lo que dijeron los codificadores humanos, cada columna lo que dijo el modelo; la diagonal es donde coinciden. Lo importante: casi todos los desacuerdos involucran a "Other", en las dos direcciones, porque Other es la categoría cajón ("todo lo demás") y la línea entre Other y un tema específico es la más difícil de trazar, en este codebook y en cualquiera. Las categorías chicas pierden más (Social Welfare y Education retienen solo 73 a 74% de sus proyectos).""",
f"""
One accuracy number hides seven different stories. This matrix shows where the disagreements are.

Each row is what the human coders said. Each column is what the model said. The cells on the diagonal are where they agree.

Let me read two cells with you. Start at the Health row. {CM['Health']['Health']} health bills were also called Health by the model. {CM['Other']['Health']} health bills were called Other by the model. Now go to the Other row. {CM['Health']['Other']} bills that the humans coded as Other were called Health by the model.

So most of the disagreements involve the Other category, in both directions. That makes sense: Other is the catch-all category, everything that is not one of the six specific topics, and the line between Other and a specific topic is the hardest line to draw. It will be the hardest line in your codebook too.

One more thing: the small categories pay for the accuracy of the big one. Social Welfare keeps only {p0(PC['Social Welfare']['recall'])} percent of its bills, Education {p0(PC['Education']['recall'])}.
"""),
(17, "Precision and recall (one breath, no stop)", "20:00",
f"""De qué va: dos palabras para los metodólogos del panel, sin detenerse. Precisión: cuando el modelo dice Salud, cuántas veces tiene razón ({p0(PC['Health']['precision'])}%). Recall: de todos los proyectos de salud según los humanos, cuántos encuentra el modelo ({p0(PC['Health']['recall'])}%). Un revisor pide estos dos números por categoría, no una sola exactitud.""",
f"""
Two more words, and then we move on. Precision: when the model says Health, how often is it right? Here, {p0(PC['Health']['precision'])} percent of the time. Recall: of all the bills the humans called Health, how many did the model find? {p0(PC['Health']['recall'])} percent. A reviewer will ask for these two numbers for each category, not for one accuracy number.
"""),
(18, "Challenge 2, part two: whose error is it?", "20:30",
"""De qué va: Challenge 2, parte dos. Tres títulos en los que el modelo y el humano no coinciden. Ellos votan quién se equivocó (el modelo, el humano, o el codebook es ambiguo) y qué cambiarían. Después de la votación das los resultados: A, el humano tiene razón y el modelo se equivocó (se fue con la palabra "defense"); arreglo: una regla en el prompt. B, el modelo se equivocó (oyó "tobacco" y dijo Health; es seguridad de producto, Other); arreglo: aclarar el codebook. C, el codebook es ambiguo (una regla del Seguro Social para cónyuges divorciados: ¿derecho de familia o bienestar social?); arreglo: decidir y dejarlo escrito. La lección: un error tiene tres dueños posibles y cada dueño tiene un arreglo distinto.""",
"""
Challenge two, part two. Whose error is it?

Here are three titles where the model and the human coder disagree. For each one I want you to vote: the model is wrong, the human is wrong, or the codebook is ambiguous. And then tell me what you would change.

Case A is a supplemental appropriation for the Defense Health Program. The humans said Health. The model said Other, and its own explanation was that this is a Defense Department bill.

Case B sets fire-safety standards for cigarettes. The humans said Other, because this is consumer product safety. The model said Health, because it heard the word tobacco.

Case C changes a Social Security rule for divorced spouses. The humans said Law and Crime, because the codebook puts family law there. The model said Social Welfare.

[Let them vote and argue for two minutes. Then give the results.]

Here is how I read them. In case A, the human is right and the model is wrong: the money goes to health care, even though the department is Defense. The fix is a rule in the prompt: a health program inside any department is still Health. In case B, the model is wrong: the fire safety of a product is consumer safety, not health. The fix is a clarification in the codebook: tobacco means use and health effects, not product safety. In case C, the codebook itself is ambiguous: a rule about divorced spouses is family law, but it is also a Social Security benefit. The fix is a decision, written down in the codebook, so that the humans and the model apply the same one.

That is the lesson of this exercise. An error can have three owners: the model, the codebook, or the human code. And each owner has a different fix: a rule in the prompt, a clarification in the codebook, or a correction of the answer key.

[Checkpoint: leave this slide at 23:00.]
"""),
(19, "The error is not the same for everyone", "23:00",
f"""De qué va: el corazón de la clase. Tomamos los mismos proyectos de ley y contamos qué porcentaje es sobre salud, separando por el partido del legislador que los presentó, dos veces: una con los códigos humanos y otra con las etiquetas del modelo. Con los humanos no hay diferencia entre partidos ({p1(BP['Democrat']['health_gold'])}% de los proyectos demócratas y {p1(BP['Republican']['health_gold'])}% de los republicanos son de salud). Con el modelo aparece una brecha de 2.5 puntos ({p1(BP['Democrat']['health_llm'])}% contra {p1(BP['Republican']['health_llm'])}%) que no existe: la creó el modelo. La exactitud del modelo es casi igual para los dos partidos ({p0(BP['Democrat']['accuracy'])}% y {p0(BP['Republican']['accuracy'])}%); lo que cambia es la dirección del error: en los proyectos demócratas cuenta salud de más, en los republicanos de menos. Por época pasa lo mismo. Conclusión: hay que validar por los grupos que vas a comparar, no solo en total.""",
f"""
Now the slide that answers the question: is eighty-eight percent good enough?

Let me explain what we are counting here. Every bill comes with the party of the legislator who sponsored it. So we can ask: among the bills sponsored by Democrats, what share is about health? And among the bills sponsored by Republicans? We can answer that question twice: once using the human codes, and once using the model's labels, on exactly the same bills.

With the human codes: among Democratic bills, {p1(BP['Democrat']['health_gold'])} percent are about health; among Republican bills, {p1(BP['Republican']['health_gold'])}. No difference between the parties.

Now the same count with the model's labels: among Democratic bills the model says {p1(BP['Democrat']['health_llm'])} percent are about health; among Republican bills, {p1(BP['Republican']['health_llm'])}. A gap of two and a half points between the parties, and that gap does not exist. The model created it.

Here is the part that matters. The model's accuracy is almost identical for the two parties: {p0(BP['Democrat']['accuracy'])} percent on Democratic bills, {p0(BP['Republican']['accuracy'])} on Republican bills. What differs is the direction of the errors. On Democratic bills the model calls too many things health; on Republican bills, too few. Equal accuracy, opposite errors, and that is what bends a comparison between groups.

Over time we see the same thing: the model over-counts health in the first period, under-counts it in the middle, and over-counts again after 2000, and its accuracy slips from {p0(BE['1973-1986']['accuracy'])} to {p0(BE['2001-2016']['accuracy'])} percent across the periods.

So the rule is: always validate by the groups you are going to compare, not just overall.
"""),
(20, "Measure the error, then subtract it", "24:30",
"""De qué va: cómo se corrige el error. La analogía: una balanza que marca de más. No la tirás: pesás a algunas personas también en una balanza buena, medís cuánto se desvía en promedio, y restás ese desvío al promedio del grupo (corregís el promedio, no a cada persona). Con las etiquetas es lo mismo: tomás una muestra al azar de títulos, los codifica una persona, en esa muestra ves la etiqueta del modelo y la verdadera, medís el error promedio del modelo y lo restás del promedio del modelo sobre todos los títulos. Eso es prediction-powered inference (PPI, de Berkeley): una línea de aritmética. La única condición: la muestra codificada a mano tiene que ser al azar.""",
"""
So what do we do about it? We measure the error, and then we subtract it.

Think of a scale that reads a little heavy. You would not throw it away. Instead, you would weigh a few people on a good scale as well, see how far off your scale is on average, and subtract that amount from the group's average weight. Notice that you correct the average, not each person.

The same move works with labels. Take a sample of titles, chosen at random, and have a person code them. On that sample you can see both the model's label and the true code, so you can measure the model's average error. Then you subtract that average error from the model's average over all the titles.

That is the formula on the screen, and it is the whole method. The corrected share equals the model's share over all the titles, minus the model's average error measured on the hand-coded sample. A label is just a one or a zero, one if the title was called Health and zero if not, so a share is simply an average.

This method is called prediction-powered inference, PPI, developed by Angelopoulos and coauthors here at Berkeley, and there is a closely related version from Egami and coauthors at Princeton. It is one line of arithmetic, and it works whatever the model's error pattern is, as long as the hand-coded sample is random. That is the one condition that matters: pick the sample at random, never the easy titles.
"""),
(21, "Three estimates of one share", "26:00",
f"""De qué va: el gráfico muestra tres maneras de estimar el porcentaje de proyectos sobre salud, para todos y por partido. Azul, "naive": contar con las etiquetas del modelo sobre los 1,650 títulos como si fueran verdad. Naranja, "solo a mano": contar en los 200 títulos codificados por una persona, ignorando el modelo. Verde, "corregida": el modelo sobre todos, menos el error medido en los 200. La marca negra es la verdad (los 1,650 con código humano, que acá conocemos y normalmente no). Lo que se aprende: la naive tiene un intervalo angosto pero está sesgada (la brecha de 2.5 puntos; en un corpus de 20,000 documentos parecería un hallazgo fuerte, porque el intervalo naive se achica con el tamaño y el sesgo no). La de solo a mano es insesgada pero ancha (200 títulos dicen poco). La corregida queda centrada en la verdad ({p1(E['Democrat']['corrected']['est'])}% contra {p1(E['Republican']['corrected']['est'])}%: la brecha desaparece) y con un intervalo un 30% más angosto que la de solo a mano, porque usa los 1,650 títulos del modelo con su error medido.""",
f"""
This chart shows three ways to estimate the same number, the share of bills about health, for all bills and then by party. Before I explain them, a prediction: which way do you think the correction will move the Democratic estimate, up or down? [Take one or two answers.]

The first estimate, in blue, is the naive one: we count the model's labels over all one thousand six hundred and fifty titles and treat them as true. For Democratic bills it says {p1(E['Democrat']['naive']['est'])} percent, for Republican bills {p1(E['Republican']['naive']['est'])}. That is the two and a half point gap we just saw. Its interval is narrow, because it ignores the model's error completely, and that is exactly the problem: a narrow interval around the wrong number. On one thousand six hundred and fifty titles the gap is marginal. But on a real corpus of twenty thousand documents, the same gap would look like a strong finding, far outside the interval, because the naive interval gets narrower as the corpus grows, while the bias does not shrink at all.

The second estimate, in orange, uses only the two hundred hand-coded titles and ignores the model. It gives {p1(E['Democrat']['human']['est'])} versus {p1(E['Republican']['human']['est'])}. This estimate is right on average, but any one sample of two hundred titles can land a point or two away from the truth, and that is what the wide interval is telling you.

The third estimate, in green, is the corrected one: the model's count over all titles, minus the error measured on the two hundred. It gives {p1(E['Democrat']['corrected']['est'])} versus {p1(E['Republican']['corrected']['est'])}. The gap is gone. And its interval is about thirty percent narrower than the hand-coded one, because we did not throw away the model's labels on the other fourteen hundred titles; we used them, with their error measured.

The black mark is the truth: what all one thousand six hundred and fifty human codes say. Normally you cannot afford that. Here it lets us check that the method works, and it does: the green estimates sit right next to the black marks.

[Optional: switch to the notebook and run the two cells that print these tables. Sixty seconds. If the screen fails, the slide has the numbers.]
"""),
(22, "Back to the opening chart", "27:30",
f"""De qué va: volvemos al gráfico del inicio. Por período, los humanos dan {p1(E['1973-1986']['truth'])}%, {p1(E['1987-2000']['truth'])}%, {p1(E['2001-2016']['truth'])}% (sube y se estabiliza); el modelo da {p1(E['1973-1986']['naive']['est'])}%, {p1(E['1987-2000']['naive']['est'])}%, {p1(E['2001-2016']['naive']['est'])}% (sigue subiendo). Después del 2000 las dos series se separan sobre los mismos proyectos: la subida después del 2000 es el instrumento, no los proyectos de ley. La corrección por período (verde) es honesta pero ancha: con solo 63 a 72 títulos a mano por período, los intervalos son de ±4 a 5 puntos: cubren la verdad pero no pueden decir si la serie se estabilizó o siguió subiendo. Y en el primer período el punto corregido queda más lejos de la verdad que el naive; con muestras tan chicas el punto se mueve, y la muestra sola (1 proyecto de salud en 63) da un intervalo de {p1(E['1973-1986']['human']['ci_low'])}% a {p1(E['1973-1986']['human']['ci_high'])}%: honesto pero inútil. Decilo vos antes de que lo digan ellos. De ahí el Challenge 3: ¿cuántos títulos a mano hacen falta?""",
f"""
Now let us go back to the opening chart. Which part of that climb was the instrument?

Here are the same three periods, estimated three ways. The human codes say {p1(E['1973-1986']['truth'])} percent, then {p1(E['1987-2000']['truth'])}, then {p1(E['2001-2016']['truth'])}: health rose and then leveled off. The model's labels say {p1(E['1973-1986']['naive']['est'])}, then {p1(E['1987-2000']['naive']['est'])}, then {p1(E['2001-2016']['naive']['est'])}: a climb that keeps going. After 2000 the two series part ways on the very same bills: the model over-counts health by two points where it had under-counted it before, and that divergence is well beyond noise. So the climb after 2000 was the instrument, not the bills.

Now look at the green intervals, the corrected estimates by period. They are honest about how much sixty-three to seventy-two hand-coded titles per period can tell you: plus or minus four to five points. They cover the truth in every period, but they are too wide to tell a plateau from a climb.

And let me say the uncomfortable part before you do. In the first period, the corrected point actually lands farther from the truth than the naive one. And the hand-coded sample alone found just one health bill in sixty-three titles, which gives an interval from {p1(E['1973-1986']['human']['ci_low'])} percent to {p1(E['1973-1986']['human']['ci_high'])} percent: honest, and useless. With samples that small, the point estimate wanders. An interval that admits its own width is the right answer; a single number would mislead you. That is the price of splitting two hundred hand-coded titles across three periods. So the natural question is: how many would you need? That is Challenge three.
"""),
(23, "Challenge 3: how many titles would you code by hand?", "28:30",
f"""De qué va: Challenge 3, el exit ticket. Escenario: 20,000 títulos; el comité quiere saber si demócratas y republicanos difieren en 3 puntos o más; un asistente codifica 100 títulos por hora; regla: para distinguir de cero una brecha de 3 puntos entre dos grupos, el intervalo de cada grupo tiene que ser de ±2 puntos o menos. Ellos leen la curva (verde: con el modelo corregido por la muestra; naranja: solo a mano) y escriben cuántos títulos por partido, cuántas horas y qué intervalo prometen. Respuesta: unos {N_C3} por partido (±{p1(PJrow[400]['ci_halfwidth'])}), {2 * N_C3} títulos, {2 * N_C3 // 100} horas; solo a mano, unos {N_H3} por partido, {2 * N_H3 // 100} horas. Con 200 por partido el intervalo es ±{p1(PJrow[200]['ci_halfwidth'])}: no alcanza. Lección: 200 al azar es el piso para estimar un solo porcentaje; comparar grupos cuesta más por grupo.""",
f"""
Challenge three, and this one is your exit ticket. How many titles would you code by hand?

Here is your project. You are using the same instrument on twenty thousand bill titles. Your committee wants to know whether Democrats and Republicans differ by three points or more in the share of health bills. You have a research assistant who codes about one hundred titles an hour. And one rule, which is on the slide: to tell a three-point gap between two groups apart from zero, the interval for each group has to be about plus or minus two points, or narrower.

The curve on the screen shows how wide the interval is for different numbers of hand-coded titles: in green with the model corrected by the sample, in orange with the hand-coded sample alone. Read the curve and write three things on the handout: how many titles per party, how many assistant hours, and the interval you would promise your committee. Three minutes.

[After three minutes, show the results.] Here is the answer you can check. On the green curve, about {N_C3} titles per party bring the interval down to plus or minus two. That is {2 * N_C3} titles in total, {2 * N_C3 // 100} hours, one day of an assistant's time. Without the model, on the orange curve, you would need about {N_H3} per party, {2 * N_H3 // 100} hours.

Now look at two hundred per party. The interval there is plus or minus {p1(PJrow[200]['ci_halfwidth'])}, so a three-point gap could not be told apart from zero. In other words, two hundred at random, which is the floor I will recommend for one share, is not enough for this question. A comparison between groups costs more per group.

Two closing remarks. Doubling the number of titles cuts the width by about thirty percent, so four times the work buys half the width. And this curve was built from this lesson's error rates; in your own project you would build it from a pilot of about one hundred hand-coded documents, and then decide. If your whole corpus is smaller than two hundred documents, say forty interviews, hand-code everything.

[Checkpoint: 32:00.]
"""),
]),
("Part C. Lesson, part two: links as measurements (32:00 to 37:00)", [
(24, "Second problem, same recipe", "32:00",
f"""De qué va: la segunda parte, record linkage con la misma receta. Febrl: dos archivos de 5,000 personas; el B es copia "sucia" del A (errores de tipeo, dígitos cambiados, direcciones distintas), y como es sintético se sabe qué registros son la misma persona (los 5,000 links verdaderos). Comparar todos contra todos serían 25 millones de pares; con "blocking" (comparar solo pares que comparten algo: apellido parecido, misma fecha de nacimiento o mismo código postal) quedan {BL['candidates']:,} pares candidatos y se pierden solo {BL['lost_links']} links verdaderos, y los {BL['lost_links']} son personas que se mudaron. Reglas simples (un punto por campo que coincide) clasifican casi todo: {BL['sure']:,} pares seguros, {BL['candidates'] - BL['sure'] - BL['band']:,} seguros que no, y {BL['band']:,} ambiguos (con {BL['band_true']} links verdaderos adentro). Ese bloque ambiguo es donde una persona o un modelo tiene que decidir.""",
f"""
Part two. A second problem, and the same recipe. The question now is: are these two records the same person?

Febrl is a test dataset for record linkage: two files of five thousand person records, where file B is a corrupted copy of file A, with typos, swapped digits and changed addresses. Because it is synthetic, we know exactly which record in A matches which record in B: all five thousand true links. We keep four fields that two administrative files would realistically share: given name, surname, date of birth, postcode. No ID number. Your version could be census manuscripts across decades, or hospital discharge records linked to death records.

If we compared every record in A with every record in B, that would be twenty-five million pairs. Instead we use blocking: we only compare pairs that share something, a similar surname, the same birth date, or the same postcode. That leaves about one hundred and thirty-eight thousand candidate pairs, and it loses only {BL['lost_links']} true links. Interestingly, all {BL['lost_links']} are people who moved: the lesson's own mechanism is already at work in this first step.

Then simple rules sort most pairs. Each pair gets one point per field that matches. Four points: surely the same person, about four thousand three hundred pairs. Zero or one point: surely different people, about one hundred and thirty-two thousand pairs. And two points: the ambiguous band, about fifteen hundred pairs, with {BL['band_true']} true links inside. Linking records is a classification problem too: the unit is the pair, the label is same person or not, and the ambiguous band is where a person or a model has to decide.
"""),
(25, "Four ways to handle the band", "33:30",
f"""De qué va: cuatro políticas para el bloque ambiguo, medidas contra la verdad. Dos palabras: precisión (de los pares que declarás "misma persona", cuántos lo son) y recall (de los links verdaderos, cuántos encontrás). Descartar el bloque: precisión 100%, recall {p0(LP['Rules, strict (drop the band)']['recall'])}% (perdés 14% de la gente). Aceptarlo entero: recall {p0(LP['Rules, lenient (keep the band)']['recall'])}%, precisión {p0(LP['Rules, lenient (keep the band)']['precision'])}% (uno de cada seis links es inventado). Que el modelo decida: precisión {p1(LP['Rules + Haiku on the band']['precision'])}%, recall {p1(LP['Rules + Haiku on the band']['recall'])}%. El modelo cuesta centavos y reemplaza un día de revisión manual, pero sigue siendo una medición: en tu proyecto no conocés la verdad, así que revisás a mano una muestra al azar del bloque (para la precisión) y una muestra de los pares justo debajo del bloque (para el recall).""",
f"""
Four ways to handle that ambiguous band, measured against the truth. Two words first. Precision: of the pairs we declare to be the same person, how many really are. Recall: of all the true links, how many we find.

If we drop the band: precision one hundred percent, recall {p0(LP['Rules, strict (drop the band)']['recall'])}. Very clean, but fourteen percent of the people are gone.

If we keep the whole band as matches: recall {p0(LP['Rules, lenient (keep the band)']['recall'])} percent, precision {p0(LP['Rules, lenient (keep the band)']['precision'])}. We find almost everyone, but one link in six is invented.

If we let the model decide inside the band: precision {p1(LP['Rules + Haiku on the band']['precision'])}, recall {p1(LP['Rules + Haiku on the band']['recall'])}. Inside the band the model finds {p0(R['band_claude-haiku-5-5']['recall_in_band'])} percent of the true pairs, and when it says yes, it is right {p0(R['band_claude-haiku-5-5']['precision_in_band'])} percent of the time. The stronger model, by the way, is not better here.

The model costs cents and replaces a day of a person checking pairs by hand. But it is still a measurement. We only know these numbers because, in this test dataset, the truth is known. In your project the truth is unknown, so you hand-review a random sample of the band, which tells you the precision, and a sample of the pairs just below the band, where the missed links hide, which tells you the recall. That is the human-in-the-loop review any automated pipeline needs before anyone trusts it.
"""),
(26, "Question: lose pairs or invent them?", "35:00",
f"""De qué va: la pregunta para el panel: ¿preferís perder pares o inventarlos? La pregunta de investigación: qué porcentaje de personas cambió de código postal (se mudó). Verdad: {p1(MV['Truth (all true links)'])}%. Descartar el bloque: {p1(MV['Rules, strict'])}%. Aceptar todo: {p1(MV['Rules, lenient'])}%. Modelo: {p1(MV['Rules + Haiku'])}%. Por qué tanta sensibilidad: el código postal es a la vez un campo de matching y la variable de resultado; los links perdidos esconden a los que se mudaron, los inventados crean mudanzas falsas. Conclusión: no hay política gratis; la elección depende del número que querés estimar; en tus datos, revisás muestras, medís los dos errores y corregís, igual que en la parte uno.""",
f"""
Now a question for you. Would you rather lose pairs or invent them?

Suppose the research question is simple: what share of people changed their postcode between the two files, in other words, what share moved? The truth, over all five thousand true links, is {p1(MV['Truth (all true links)'])} percent.

If we drop the ambiguous band, we estimate {p1(MV['Rules, strict'])}. If we keep the whole band, {p1(MV['Rules, lenient'])}. If we let the model decide, {p1(MV['Rules + Haiku'])}.

Why is this number so sensitive to the policy? Because the postcode is both a field we use for matching and the outcome we want to measure. A pair where the person moved is exactly the kind of pair the rules find hard, so dropping the band hides the movers, and keeping everything creates movers that do not exist.

[Take their answers.] There is no free policy. The right choice depends on the number you want to know: for a mobility rate, missed links push it down and invented links push it up. Here we can see all of this because the moves were planted in the data. In your own data the truth is unknown, so you review random samples of the band and of the near-misses, you measure both error rates, and you correct. Exactly like part one.
"""),
]),
("Part D. Close (37:00 to 40:00)", [
(27, "What we know works", "37:00",
"""De qué va: el cierre: la lista de prácticas que un estudiante puede llevar a su apéndice metodológico. Tres sobre el instrumento (codebook dentro del prompt; formato de respuesta fijo y misma configuración en todas las llamadas; registrar el instrumento), dos sobre validación (muestra aleatoria reportada por categoría y por grupo; estabilidad: correr dos veces y reportar el acuerdo) y una sobre estimación (corregir).""",
"""
Let me close with what we know works, as of 2026. This is a list a student can take straight to a methods appendix.

Three things about the instrument. Put the codebook in the prompt: in this lesson the definitions and the rules for the hard cases were worth fourteen points of accuracy; the examples added nothing we could measure; the stronger model, one and a half points. Use a fixed answer format and the same settings for every call, so that a label today and a label next month come from the same instrument. And log the instrument: model and version, prompt version, date, tokens.

Two things about validation. A random validation sample, with the results reported by category and by the groups you compare. And stability: the same title can get a different label on a second run, so run a random sample twice and report how often the two runs agree.

And one thing about estimation: correct it. Never treat AI labels as the truth in a regression or in a trend.
"""),
(28, "Report, privacy, cost (thirty seconds)", "38:00",
"""De qué va: tres tarjetas en treinta segundos: qué reportar; privacidad (el texto enviado a un modelo por internet sale de tu máquina; con datos restringidos, como un proyecto del FSRDC o datos bajo un acuerdo de uso, nada sale del entorno seguro: se valida adentro y la revisión de divulgación decide qué celdas de la tabla de errores se pueden publicar); y costo.""",
"""
Three cards, thirty seconds. What to report: the codebook and the prompt, word for word; the model and version; the validation sample; the metrics by category and by group; how the estimate was corrected; and where the code lives.

Privacy: text sent to a model over the internet leaves your machine. Interview transcripts, and anything under an IRB protocol, need a data use review first. With restricted data, an FSRDC project or a data use agreement, nothing leaves the secure environment: you validate inside it, and disclosure review decides which cells of your error table can come out.

Cost: this whole run was under twenty dollars and about an hour of computer time. The two hundred hand-coded titles took one afternoon, and that afternoon is what makes the dollars usable.
"""),
(29, "Key points and take-home", "38:30",
"""De qué va: los cinco puntos clave (el resumen de los cinco objetivos) y el take-home challenge, que conecta la clase con la consultoría del D-Lab. Al final recogés los handouts con la respuesta del Challenge 3.""",
"""
Five key points, and then I stop.

One: an AI label or an AI link is a measurement with error. The prompt is the instrument, so lock it before you score anything.

Two: validate on a random hand-coded sample, by category and by the groups you will compare. Accuracy alone can hide a finding that is not there.

Three: high accuracy does not make an estimate unbiased. Measure the error and subtract it; it is one line of arithmetic.

Four: how many documents to hand-code is a design decision you read off a curve. Two hundred at random is the floor for one share, and a comparison between groups costs more per group.

Five: the same recipe handles record linkage, and the downstream number moves with the linkage policy, so report the policy together with the number.

And a take-home challenge: take a thousand documents from your own project, hand-code two hundred chosen at random, run the notebook's three estimates on a share you care about, and bring the corrected interval to a D-Lab consultation. If you have fewer than two hundred documents, code them all.

[Collect the handouts with the Challenge three answers.]
"""),
(30, "Thank you", "39:30",
"""De qué va: cerrás y abrís las preguntas.""",
"""
That is the lesson. Thank you. I am happy to take your questions about it.
"""),
]),
]

CUTS = [
    "Dicho completo, este guion dura unos 48 minutos (5,800 palabras más los silencios de los ejercicios). Para entrar en 40 hay dos palancas. La primera: los ejercicios en 2 minutos cada uno en vez de 3, 2 y 3 (ahorra 3 minutos). La segunda: los párrafos de abajo se pueden soltar sin perder el hilo (ahorran unos 6 minutos). Elegí vos cuáles.",
    "Párrafos que se pueden soltar: slide 10, el último ('your version'); slide 11, el último ('Why the word instrument'); slide 13, la tercera lección (la clave de respuestas también tiene errores); slide 19, el párrafo de 'Over time'; slide 21, la oración sobre los 20,000 documentos y el párrafo de la marca negra; slide 22, el párrafo de 'the uncomfortable part'; slide 23, los 'two closing remarks'; slide 24, la frase 'Your version'; slide 25, la frase sobre el modelo más fuerte; slide 27, decir solo la primera oración de cada bloque.",
    "Checkpoint a las 23:00, saliendo del Challenge 2b (slide 18): si pasaste de 24:30, decí el slide 20 en dos oraciones (la balanza y la fórmula) y en el slide 21 leé solo los párrafos del azul y del verde.",
    "Checkpoint a las 32:00, saliendo del Challenge 3 (slide 23): si pasaste de 33:30, saltá el slide 25 y decí sus tres resultados en el slide 26 (descartar: 100 y 86; aceptar: 85 y 98; modelo: 99.5 y 94), y saltá el slide 28.",
    "Nunca saltes un ejercicio. Si todo lo demás falla, el Challenge 1 pasa a dos minutos y el Challenge 3 a tres.",
]

PRONUNCIATION = [
    ("kappa", "KAP-uh"),
    ("Febrl", "FEB-rul (say it once, then call it 'the test dataset')"),
    ("Angelopoulos", "an-jel-OP-oo-los"),
    ("Egami", "eh-GAH-mee"),
    ("Desai", "deh-SIGH"),
    ("Fellegi-Sunter", "FELL-eh-ghee SUN-ter (only if someone asks about probabilistic matching)"),
    ("adjudicate", "uh-JOO-dih-kate; or simply say 'decide'"),
    ("enclave", "EN-klayv; the script says 'secure environment' instead"),
    ("naive", "nah-EEV"),
    ("Haiku, Sonnet", "HY-koo, SON-net"),
]

# ---------------------------------------------------------------------------
# Build the document
# ---------------------------------------------------------------------------
doc = Document()
for sct in doc.sections:
    sct.top_margin = sct.bottom_margin = Inches(0.9)
    sct.left_margin = sct.right_margin = Inches(1.0)
st = doc.styles["Normal"]
st.font.name = "Calibri"
st.font.size = Pt(12)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
st.paragraph_format.space_after = Pt(8)
st.paragraph_format.line_spacing = 1.15


def heading(text, size=16, color=INK, before=12, after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.font.color.rgb = color
    return p


def body(text):
    """A paragraph of script; [bracketed] directions are set in italics and grey."""
    p = doc.add_paragraph()
    pos = 0
    for m in re.finditer(r"\[[^\]]+\]", text):
        if m.start() > pos:
            p.add_run(text[pos:m.start()])
        r = p.add_run(m.group(0))
        r.italic = True
        r.font.color.rgb = MUTED
        pos = m.end()
    if pos < len(text):
        p.add_run(text[pos:])
    return p


def guide_box(text):
    """The Spanish note on what the slide is about, in a shaded one-cell table."""
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    tblPr = t._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    c = t.cell(0, 0)
    c.width = Inches(6.5)
    tcPr = c._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), GUIDE_FILL)
    tcPr.append(shd)
    c.text = ""
    p = c.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    r.font.size = Pt(10.5)
    r.font.color.rgb = INK
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


heading("Guía y guion de la presentación: Validate Before You Estimate", size=20, before=0)
p = doc.add_paragraph()
r = p.add_run("D-Lab mock lesson, jueves 5 de noviembre de 2026, 2:40 PM. Overview más lección, 40 minutos, después preguntas.")
r.font.color.rgb = MUTED

heading("Cómo usar este documento", size=13, color=BLUE)
for t in [
    "Cada slide tiene dos partes. El recuadro sombreado explica en español de qué va el slide y cuál es el punto que tiene que quedar claro; es para vos, no se lee en voz alta. Debajo va el guion en inglés, lo que decís, escrito como una conversación que explica el razonamiento paso a paso.",
    "La hora junto a cada título es el minuto en que ese slide debería empezar. Las palabras entre [corchetes] son indicaciones, no se leen.",
    "No hace falta decir cada oración. Si perdés el hilo, leé el primer bullet del slide y seguí desde ahí. Los números están en pantalla; podés señalarlos en vez de leerlos todos.",
    "Leelo en voz alta dos veces, la segunda con cronómetro. Marcá las oraciones que querés decir sí o sí y las que podés soltar si el reloj se pone en contra; al final del documento están los cortes sugeridos.",
]:
    body(t)

total_words = 0
for part_title, slides in SCRIPT:
    heading(part_title, size=14, color=BLUE, before=18)
    for n, title, clock, guide, text in slides:
        heading(f"Slide {n}  |  {title}  |  {clock}", size=12, before=10, after=4)
        guide_box(guide.strip())
        for para in [t.strip() for t in text.strip().split("\n\n") if t.strip()]:
            body(para)
            total_words += len(re.sub(r"\[[^\]]+\]", "", para).split())

heading("Si vas atrasado", size=14, color=BLUE, before=18)
for t in CUTS:
    body(t)

heading("Pronunciación", size=14, color=BLUE, before=18)
for word, how in PRONUNCIATION:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(word + ": ")
    r.bold = True
    p.add_run(how)

out = PROJECT / "docs/Presentation_Script.docx"
doc.save(out)
print(f"script written: {out}  ({total_words} spoken words)")
