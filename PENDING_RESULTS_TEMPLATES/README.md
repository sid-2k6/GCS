# Pending result files — schemas only

**Every CSV in this folder is empty of results on purpose.** The numbers must come from an executed run of
`PsyAdapt_Reviewer_Revision.ipynb`; inventing them is what the reviewer objected to in the first place. The
column names match what the notebook already writes, so filling them is mechanical.

Run `harvest_pipeline_outputs.py --run-dir <pipeline output root> --outdir ./filled` to populate them from
real per-response data. It prints, for each item, either `[OK]` with the manuscript location it unblocks or
`[MISSING]` with the path it expected — anything still missing must keep its `PENDING` / `UNVERIFIED` marker
in the paper.

## Files and what each one unblocks

| File | Produced by | Manuscript location it closes |
|---|---|---|
| `paired_contrasts_all_metrics.csv` | §20 → `statistics/paired_contrasts.csv` | §4.10 `[PENDING]`: paired tests for ROUGE-L, input relevance, grounding, defects, latency. **The defect contrasts matter most** — that is where the real component effect sits |
| `friedman_omnibus.csv` | §20 → `statistics/friedman_omnibus.csv` | supporting sentence in §4.10 |
| `length_controlled_contrasts.csv` | §22c → `evaluation/length_controlled_contrasts.csv` | §4.12 `[PENDING]`: whether the A7−A2 gap survives within matched length strata |
| `confusion_matrices_long.csv` | recomputed from `baselines/predictions/*.csv` | reviewer R09, which asked for confusion matrices; §4.9 currently reports aggregate metrics only |
| `judge_pairwise_winrates.csv` | §18a → `evaluation/pairwise_winrates.csv` | §4.14: the bias-corrected win rates and position-bias rate, none of which is in the paper yet |
| `chunk_granularity_contrast.csv` | §20 → `evaluation/chunk_granularity_contrast.csv` | §3.2.5, which asserts the re-chunking was necessary without showing numbers |
| `chunk_granularity_stats.csv` | §10 → `retrieval/chunk_granularity_stats.csv` | the three `[UNVERIFIED chunk counts]` markers (Table 2, Table 6, §3.2.5). Note the conflict to resolve: the old paper says **81** chunks, the notebook's §10 diagnostic says the v2 KB had **102** chunks with median 3,029 characters |
| `sota_A0_bootstrap_ci.csv` | recomputed from `evaluation/automatic_metrics_per_response.csv` | the "CI not reported" cell for the bare Qwen arm in Table 9 |
| `figure6_defects_by_topic.csv` | recomputed from `defect_counts.csv` + `benchmark_component_features.csv` | the `PENDING` note in the Figure 6 caption |

## Provenance files — fill these before submission

These three record gaps I could not resolve from the notebook as supplied. They are not optional
bookkeeping: each concerns a number that is already printed in the manuscript.

| File | Question to answer |
|---|---|
| `statistical_contrast_provenance_TO_CONFIRM.csv` | The notebook's `CONTRASTS` list produces `A1-A0, A2-A1, A3-A2, A4-A2, A5-A2, A6-A2, A7-A2, A7-A6, A7-A0`. The supplied `statistical_contrasts_bertscore.csv` also contains **A7-A1, A7-A3, A7-A4, A7-A5**, which that list does not produce. Table 12, Figure 9, the abstract and the conclusion all rest on these, so record how they were produced |
| `memory_experiment_provenance_TO_CONFIRM.csv` | `memory_experiment.csv` reports BERTScore-F1 for A2-M vs A5-M, but §20's memory contrast is coded against judge scores (`context_continuity`, `overall`) and §17 computes automatic metrics only for the main benchmark. Record which cell produced 0.618 / 0.584 |
| `defect_category_mapping_TO_CONFIRM.csv` | The Table 16 labels match neither the six judge categories in §18b nor exactly the regex screen names in §17. Map each label to its pipeline field, or rename the labels to match the pipeline |

## After filling them

1. Insert the values into the manuscript and delete the corresponding marker — search the .docx for
   `PENDING`, `UNVERIFIED`, `REQUIRES HUMAN ACTION` (8 hits).
2. Highlight each newly inserted value in yellow, matching the existing revision marking.
3. Re-check the claim wording: if a newly available paired test is not significant after Holm correction,
   the surrounding sentence must say so, exactly as §4.4, §4.5, §4.8 and §4.10 already do for A7 vs A2.
4. Move the finalized CSV into `UPDATED_TABLES/` and add a row to `CHANGE_LOG.csv` and to
   `FINAL_CONSISTENCY_REPORT.docx`.
