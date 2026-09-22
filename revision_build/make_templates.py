"""Create the empty annotation sheet and the header-only templates for the pending result files."""
import csv
import os

HE = "/projects/sandbox/GCS/HUMAN_EVALUATION"
PT = "/projects/sandbox/GCS/PENDING_RESULTS_TEMPLATES"
os.makedirs(HE, exist_ok=True)
os.makedirs(PT, exist_ok=True)

CRITERIA = ["relevance", "empathy", "personalization", "helpfulness", "safety", "overall"]
N_ITEMS = 100


def write(path, header, rows=()):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    print("wrote", os.path.relpath(path, "/projects/sandbox/GCS"), f"({len(list(rows))} data rows)")


# ---------------------------------------------------------------- human evaluation
sheet_header = ["item_id", "user_message", "reply_A", "reply_B"] + CRITERIA + ["comments"]
sheet_rows = [[f"H{i:04d}"] + [""] * (len(sheet_header) - 1) for i in range(1, N_ITEMS + 1)]
write(f"{HE}/human_pairwise_sheet_BLINDED_EMPTY.csv", sheet_header, sheet_rows)

key_header = ["item_id", "sample_id", "contrast", "first_config", "second_config"]
key_rows = [[f"H{i:04d}", "", "", "", ""] for i in range(1, N_ITEMS + 1)]
write(f"{HE}/human_pairwise_key_DO_NOT_SHARE_EMPTY.csv", key_header, key_rows)

write(f"{HE}/table17_human_evaluation_TEMPLATE.csv",
      ["Contrast", "Criterion", "Judgements", "Wins (first config)", "Wins (second config)", "Ties",
       "Win rate (of decided)", "Exact binomial p", "Holm-corrected p", "Reading"])

write(f"{HE}/human_agreement_TEMPLATE.csv",
      ["criterion", "annotator_a", "annotator_b", "n_items", "raw_agreement", "cohens_kappa"])

write(f"{HE}/annotator_register.csv",
      ["annotator_id", "file_returned", "clinical_qualification (yes/no/detail)", "items_completed",
       "date_returned", "notes"],
      [["", "", "", "", "", ""] for _ in range(4)])

# ---------------------------------------------------------------- pending result templates
# 1. paired contrasts for every automatic metric (unblocks Section 4.10 PENDING note)
write(f"{PT}/paired_contrasts_all_metrics.csv",
      ["metric", "contrast", "n_pairs", "mean_a", "mean_b", "mean_diff", "ci95_low", "ci95_high",
       "cohens_dz", "rank_biserial", "p_value", "p_holm", "test", "identical_output_rate"])

# 2. Friedman omnibus
write(f"{PT}/friedman_omnibus.csv",
      ["metric", "n_complete_samples", "friedman_chi2", "p_value", "missing_samples"])

# 3. length-controlled contrasts (unblocks Section 4.12 PENDING note)
write(f"{PT}/length_controlled_contrasts.csv",
      ["quartile", "metric", "contrast", "mean_diff_conditional_on_length", "n"])

# 4. confusion matrices (completes reviewer R09)
write(f"{PT}/confusion_matrices_long.csv",
      ["component", "model", "true_label", "predicted_label", "count"])

# 5. LLM-judge pairwise win rates incl. position bias
write(f"{PT}/judge_pairwise_winrates.csv",
      ["contrast", "criterion", "n", "wins_a", "wins_b", "ties", "win_rate_raw",
       "win_rate_bias_corrected", "order_flip_rate", "p_value", "p_holm"])

# 6. coarse vs fine retrieval contrast (evidence for the re-chunking claim in 3.2.5)
write(f"{PT}/chunk_granularity_contrast.csv",
      ["metric", "n_pairs", "mean_fine", "mean_coarse", "mean_diff_fine_minus_coarse",
       "ci95_low", "ci95_high", "cohens_dz", "p_value", "p_holm"])

# 7. chunk counts (removes the UNVERIFIED markers in Table 2, Table 6, Section 3.2.5)
write(f"{PT}/chunk_granularity_stats.csv",
      ["granularity", "n_chunks", "median_chars", "p90_chars", "total_chars",
       "pct_within_encoder_window_1000_chars"],
      [["coarse", "", "", "", "", ""], ["fine", "", "", "", "", ""]])

# 8. missing bootstrap CI for the bare Qwen arm in Table 9
write(f"{PT}/sota_A0_bootstrap_ci.csv",
      ["arm", "metric", "n", "mean", "ci95_low", "ci95_high"],
      [["A0_Qwen2.5-3B (bare)", "bertscore_f1", "300", "0.589", "", ""]])

# 9. Figure 6 source data (topic x configuration defect means)
write(f"{PT}/figure6_defects_by_topic.csv", ["config_id", "topic", "n_responses", "mean_defects"])

# 10. defect-category provenance: paper labels are verified, the pipeline mapping is not
write(f"{PT}/defect_category_mapping_TO_CONFIRM.csv",
      ["paper_label_table16", "count_per_response_verified", "produced_by (rule_screen | llm_judge | both)",
       "pipeline_field_name", "status"],
      [["Unsupported medical claim", "0.21", "", "", "TO CONFIRM"],
       ["Overstated certainty", "0.14", "", "", "TO CONFIRM"],
       ["Prescriptive advice without disclaimer", "0.12", "", "", "TO CONFIRM"],
       ["Referral missing (high-distress input)", "0.09", "", "", "TO CONFIRM"],
       ["Diagnostic overreach", "0.08", "", "", "TO CONFIRM"],
       ["Inappropriate boundary", "0.05", "", "", "TO CONFIRM"],
       ["Total defects per response", "0.69", "", "", "TO CONFIRM"]])

# 11. provenance ledger for the contrasts that the notebook's CONTRASTS list does not produce
write(f"{PT}/statistical_contrast_provenance_TO_CONFIRM.csv",
      ["contrast_in_table12", "in_notebook_CONTRASTS_list", "how_it_was_produced", "status"],
      [["A7-A0", "yes", "", "OK"],
       ["A7-A1", "no", "", "TO CONFIRM"],
       ["A7-A2", "yes", "", "OK"],
       ["A7-A3", "no", "", "TO CONFIRM"],
       ["A7-A4", "no", "", "TO CONFIRM"],
       ["A7-A5", "no", "", "TO CONFIRM"],
       ["A7-A6", "yes", "", "OK"]])

# 12. memory experiment provenance
write(f"{PT}/memory_experiment_provenance_TO_CONFIRM.csv",
      ["value_in_table13", "supplied_value", "notebook_cell_that_computes_it", "status"],
      [["mean BERTScore-F1 with memory (A2-M)", "0.618", "", "TO CONFIRM"],
       ["mean BERTScore-F1 without memory (A5-M)", "0.584", "", "TO CONFIRM"],
       ["paired difference / CI / d_z / Wilcoxon p", "0.034 / [0.001, 0.066] / 0.34 / 0.048", "",
        "TO CONFIRM"]])
