# Human evaluation — ready-to-run package

This closes the reviewer's central demand ("a blinded human rating on a subset — empathy, helpfulness,
safety — ideally by someone with clinical training"). The protocol is already described in **Section 4.14**
of the revised manuscript; nothing there claims results, and nothing here contains ratings.

Criteria match the automated pairwise protocol exactly, so human and machine judgments stay comparable:
`relevance`, `empathy`, `personalization`, `helpfulness`, `safety`, `overall`.

## Files

| File | Purpose |
|---|---|
| `ANNOTATION_GUIDELINES.md` | Give this to every annotator. Defines each criterion, the tie policy, the rule that length and polish are not merits, the safety criterion in detail, and a distress note |
| `human_pairwise_sheet_BLINDED_EMPTY.csv` | The sheet schema with 100 pre-assigned `item_id`s (`H0001`–`H0100`) and empty content and rating columns. Usable immediately if you paste the message and two replies in by hand |
| `human_pairwise_key_DO_NOT_SHARE_EMPTY.csv` | Matching private key schema: `item_id → sample_id, contrast, first_config, second_config`. **Never send this to annotators** |
| `build_human_eval_sheet.py` | Preferred route: builds both files automatically from real generations, with order randomization and a fixed seed |
| `aggregate_human_eval.py` | Scores the returns and writes the paper-ready table |
| `table17_human_evaluation_TEMPLATE.csv` | Column layout of the table that goes into Section 4.14 |
| `human_agreement_TEMPLATE.csv` | Column layout for inter-annotator agreement |
| `annotator_register.csv` | Record who annotated what and whether they hold a clinical qualification, so clinician ratings can be reported separately |

## Procedure

**1. Build the sheet** (in the session where the pipeline outputs are readable):

```bash
python build_human_eval_sheet.py \
    --generations <run_v3>/generations/generations_final.csv \
    --benchmark   <path>/benchmark_300.csv \
    --n-items 100 --seed 43 --n-annotators 3 --outdir .
```

100 items are drawn across three contrasts — A7 vs A2, A2 vs A1, A2 vs A6 — which is the top of the
reviewer's 50–100 range. Items where the two configurations produced identical text are dropped, because
they carry no information. A and B are order-randomized per item and the identities go only to the key.

**2. Annotate.** Send each annotator their `human_pairwise_sheet_BLINDED_annotatorN.csv` plus
`ANNOTATION_GUIDELINES.md`. **Use at least two annotators** — with one, no agreement statistic exists and
the reviewer's concern is not answered. Returns come back as `human_pairwise_<initials>.csv` with `item_id`
and row order unchanged.

**3. Score:**

```bash
python aggregate_human_eval.py --dir . \
    [--judge <run_v3>/evaluation/pairwise_judgments.jsonl]
```

Outputs: `human_pairwise_winrates.csv`, `human_agreement.csv` (raw agreement + Cohen's kappa per pair per
criterion), `table17_human_evaluation.csv`, and `human_vs_judge_agreement.csv` when the judge file is given.
Blank cells are ignored, and any contrast × criterion with fewer than 8 non-tied judgments is reported as
"Not tested (n too small)" rather than given a p-value.

**4. Put it in the paper.** Add the table as **Table 17** in Section 4.14, replace the
`[REQUIRES HUMAN ACTION]` paragraph with the results, highlight the new text in yellow, and update
`CHANGE_LOG.csv`, `REVIEWER_RESPONSE_MATRIX.docx` (R07, R34) and `FINAL_CONSISTENCY_REPORT.docx`.

## Reporting rules — these keep the revision honest

- Report the win rate **and** the exact binomial p **and** the Holm-corrected p. A win rate of 0.55 on 60
  decided judgments is not a result on its own.
- Report inter-annotator agreement even when it is poor. Low agreement on empathy is itself a finding, and
  concealing it would repeat the problem the reviewer raised.
- Human ratings measure perceived quality of text. They do **not** establish clinical safety or therapeutic
  effect, and Sections 3.10, 4.13, 4.14 and 4.16 must keep saying so.
- If annotators are not clinically trained, say that plainly and keep the clinical-validation gap in the
  limitations.
- If the human result contradicts the automatic result — for example if annotators prefer A2 over A7 — report
  the contradiction. It is consistent with Section 4.10, where A7 vs A2 is already not statistically
  significant.
