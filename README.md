# D-Lab Validate Before You Estimate Workshop

[![Datahub](https://img.shields.io/badge/launch-datahub-blue)](https://datahub.berkeley.edu/hub/user-redirect/git-pull?repo=https%3A%2F%2Fgithub.com%2Ffelipehrv%2Fvalidate-before-you-estimate&urlpath=lab%2Ftree%2Fvalidate-before-you-estimate%2F&branch=main)
[![Binder](https://mybinder.org/badge_logo.svg)](https://mybinder.org/v2/gh/felipehrv/validate-before-you-estimate/main?urlpath=lab)
[![License: CC BY 4.0](https://img.shields.io/badge/License-CC_BY_4.0-lightgrey.svg)](https://creativecommons.org/licenses/by/4.0/)

Using AI to label text and link records in social science research. A 45-minute lesson for graduate students, prepared by Felipe H. Rodriguez Villalta for the D-Lab, UC Berkeley (mock lesson, November 2026). The method comes from Berkeley (prediction-powered inference) and Princeton (design-based supervised learning); this lesson is a short, no-code version with a hand-coding budget decision, and the same recipe applied to record linkage.

## Prerequisites

No prerequisites for the room exercises: hand-coding, voting and reading a curve need no code at all; a first statistics course helps in the estimation block. We recommend [Python Fundamentals](https://github.com/dlab-berkeley/Python-Fundamentals) to run the notebooks afterward, and the lesson builds on [Python APIs for Large Language Models](https://github.com/dlab-berkeley/Python-APIs-for-Large-Language-Models) and [Prompt Engineering](https://github.com/dlab-berkeley/prompt-engineering). No experience with AI tools is required.

Check out D-Lab's [Workshop Catalog](https://dlab-berkeley.github.io/dlab-workshops/) to browse all workshops, see what's running now, and review prerequisites.

## Workshop Goals

In this lesson, we treat every label an AI produces as a measurement with error. First, we turn a codebook into three classification prompts and compare the labels of a cheap and a strong model with the codes of trained human coders. Then, we measure the error (accuracy, kappa, precision and recall, by class and by group) and show that high accuracy does not protect a comparison between groups. Finally, we correct the estimate with a small hand-coded sample (prediction-powered inference), decide how many documents to hand-code, and apply the same recipe to record linkage, where the AI adjudicates ambiguous pairs.

## Learning Objectives

After this lesson, you will be able to:

- Explain why a classification prompt is a measurement instrument, and what changes when the codebook goes into it.
- Design a validation sample and read agreement metrics (accuracy, kappa, precision and recall) as measurement error.
- Explain why high accuracy does not guarantee an unbiased estimate or a valid comparison between groups.
- Read a bias-corrected estimate that combines AI labels with a small hand-coded sample, and choose how many documents to code by hand from an interval curve. The notebook lets you run the three estimators yourself.
- Apply the same recipe to record linkage: rules for the easy pairs, an AI adjudicator for the ambiguous band, a reviewed sample to measure the error.

This lesson does not cover the following:

- Calling LLM APIs from Python and structured outputs. These are covered in [Python APIs for Large Language Models](https://github.com/dlab-berkeley/Python-APIs-for-Large-Language-Models).
- Writing and improving prompts. This is covered in [Prompt Engineering](https://github.com/dlab-berkeley/prompt-engineering).
- Auditing an agent's qualitative codes passage by passage. This is covered in [Qualitative Coding with AI Tools](https://github.com/dlab-berkeley/Qualitative-Coding-AI-Tools).
- Model internals, fine-tuning, inductive coding, and agent pipelines (see the D-Lab's Agentic AI for Knowledge Work workshop).

## Workshop Structure

The lesson has two notebooks in the `lessons` folder, each with a matching notebook in `solutions` (`*_Solutions.ipynb`) with the challenges filled in.

1. `1_Labels_as_Measurements.ipynb`: codebook to prompt, validation, differential error, three estimates of one share, how many to hand-code.
2. `2_Links_as_Measurements.ipynb`: rules for the easy pairs, the model on the ambiguous band, precision and recall, the downstream estimate.

The in-room exercises (the handout's Challenges 1 to 3 and the linkage question) are hand-coding, betting and reading a curve; the notebook challenges are their coding counterparts and are numbered on their own.

Every model call was run beforehand and cached in `outputs/`, so the notebooks run without an API key. `outputs/dry_run_example.json` shows exactly what one call sends and what comes back. To reproduce the labels you need a key in `secrets/api_key.txt` and about 20 dollars of credit:

```
python3 scripts/01_prepare_sample.py            # sample 1,650 titles from the CAP bills file
python3 scripts/02_classify_titles.py --models claude-haiku-5-5
python3 scripts/02_classify_titles.py --models claude-sonnet-5-5 --prompts prompt_3_codebook_examples
python3 scripts/04_linkage_rules.py             # rules and the ambiguous band (Febrl, no key needed)
python3 scripts/05_adjudicate_pairs.py --models claude-haiku-5-5 claude-sonnet-5-5
python3 scripts/03_analysis.py                  # every number and figure in the slides
```

## Data

- `data/processed/bills_sample.csv`: 1,650 bill titles (75 per Congress, 93rd to 114th, 1973 to 2016) with the human major-topic code collapsed to seven labels, from the Comparative Agendas Project U.S. congressional bills dataset (Adler and Wilkerson; comparativeagendas.net). The raw file (116 MB) is not tracked; download it from the CAP site into `data/raw/`.
- `data/processed/example_pool.csv`: 42 titles outside the evaluation sample, used for the hand-coding exercise. The 14 few-shot examples in `prompts/prompt_3_codebook_examples.txt` come from an earlier draw of this pool; none of the 14 is in the evaluation sample (checked by exact title).
- `outputs/llm_labels_clean.csv`: 6,600 labels (three prompts with claude-haiku-5-5, one with claude-sonnet-5-5), each with the model's one-sentence reason, token counts and timestamp.
- `outputs/linkage_scores.csv`, `outputs/linkage_ambiguous_pairs.csv`, `outputs/llm_pairs_clean.csv`: the Febrl 4 candidate pairs, the ambiguous band, and the models' decisions on it. Febrl ships with the `recordlinkage` package.
- `outputs/results.json`: every number quoted in the slides.
- `outputs/dry_run_example.json`: one real call from the cached run, exactly as sent and received.

## Installation Instructions

Anaconda is the easiest way to get Python and Jupyter. Then:

1. Download or clone this repository.
2. In a terminal inside the repository folder, run `pip install -r requirements.txt`.
3. Open `lessons/1_Labels_as_Measurements.ipynb` in Jupyter Lab and run the cells with Shift + Enter.

## Run the Code

1. Open Jupyter Lab (Anaconda Navigator, or `jupyter lab` in a terminal inside the repository folder).
2. Open `lessons/1_Labels_as_Measurements.ipynb` and run the cells one by one with Shift + Enter; the cached outputs make the notebooks self-contained.
3. Open `lessons/2_Links_as_Measurements.ipynb` the same way. The filled-in versions are in `solutions/`.

## Is Python Not Working on Your Computer?

Use the DataHub button at the top of this page (a CalNet ID is all you need) or the Binder button; both open the notebooks in a browser with everything installed.

## Additional Resources

- Angelopoulos, Bates, Fannjiang, Jordan and Zrnic (2023), "Prediction-powered inference", *Science* 382: the one-line correction used in the lesson.
- Egami, Hinck, Stewart and Wei (2023, 2024), design-based supervised learning: valid downstream inference with LLM annotations in the social sciences.
- Desai, Card and Jacobs (2026), "Validating LLMs in social science: Epistemic threats and emerging norms" (arXiv:2607.07915): the review of validation practice in eight flagship journals quoted in the lesson.
- Comparative Agendas Project master codebook: the definitions behind the seven labels.
- D-Lab workshops: [Python APIs for Large Language Models](https://github.com/dlab-berkeley/Python-APIs-for-Large-Language-Models); [Prompt Engineering](https://github.com/dlab-berkeley/prompt-engineering); [Qualitative Coding with AI Tools](https://github.com/dlab-berkeley/Qualitative-Coding-AI-Tools); AI Pulse: LLMs for Qualitative Work.

## About the UC Berkeley D-Lab

D-Lab works with Berkeley faculty, research staff, and students to advance data-intensive social science and humanities research. Our goal at D-Lab is to provide practical training, staff support, resources, and space to enable you to use data science in your own research applications. Our services cater to all skill levels and no programming, statistical, or computer science backgrounds are necessary.

## Contributors

- Felipe H. Rodriguez Villalta

## License

The teaching materials (slides, handout, notebooks and text) are released under [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/). The bill titles and topic codes come from the Comparative Agendas Project and keep its terms of use; the Febrl records ship with the `recordlinkage` package.
