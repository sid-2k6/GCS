"""Produce REVIEWER_RESPONSE_MATRIX.docx, FINAL_CONSISTENCY_REPORT.docx and UPDATED_TABLES/."""
import csv
import os
import shutil
import sys

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Twips

sys.path.insert(0, "/projects/sandbox/GCS/_work")
import revise_doc as R  # noqa: E402  (payload tables live here)

OUT = "/projects/sandbox/GCS"
TSRC = "/projects/sandbox/GCS/_work/tables"


# ------------------------------------------------------------------ doc helpers
def new_doc(landscape=True):
    d = Document()
    s = d.sections[0]
    if landscape:
        s.orientation = WD_ORIENT.LANDSCAPE
        s.page_width, s.page_height = s.page_height, s.page_width
    for name in ("Normal",):
        st = d.styles[name]
        st.font.name = "Times New Roman"
        st.font.size = Pt(11)
    return d


def heading(d, text, size=14, align=None):
    p = d.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.font.name = "Times New Roman"
    if align:
        p.alignment = align
    return p


def body(d, text, size=11):
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.name = "Times New Roman"
    return p


def table(d, header, rows, widths=None, size=9):
    t = d.add_table(rows=1, cols=len(header))
    t.style = d.styles["Table Grid"]
    for j, h in enumerate(header):
        c = t.rows[0].cells[j]
        c.text = ""
        r = c.paragraphs[0].add_run(h)
        r.bold = True
        r.font.size = Pt(size)
        r.font.name = "Times New Roman"
    for row in rows:
        cells = t.add_row().cells
        for j, v in enumerate(row):
            cells[j].text = ""
            r = cells[j].paragraphs[0].add_run(str(v))
            r.font.size = Pt(size)
            r.font.name = "Times New Roman"
    if widths:
        t.autofit = False
        for row in t.rows:
            for cell, w in zip(row.cells, widths):
                cell.width = Twips(w)
    return t


# ------------------------------------------------------------------ 1. reviewer matrix
MATRIX = [
    ("R02", "Statistical",
     "The full system beats A2 by 0.001 on BERTScore and grounding on 300 samples from a single run, with "
     "no confidence intervals or significance tests; A4/A5 sit within 0.002 of A7; the repeated claim that "
     "A7 'performs better' is unsupported. Bootstrap CIs and a paired test are required and all claims "
     "must be rewritten to match.",
     "Added a paired analysis over the 300 items (Wilcoxon signed-rank, 5,000-resample bootstrap CIs, "
     "Cohen's d_z, Holm correction) and rewrote every comparative claim. A7 vs A2 (+0.005, CI [0.000, "
     "0.010], Holm p = 0.287) and A7 vs A5 are now reported as numerically higher but not statistically "
     "significant; the abstract, 4.4, 4.5, 4.8 and the conclusion state that retrieval plus profiling "
     "accounts for most of the gain and that emotion/strategy mainly reduce defects.",
     "Abstract; 4.1; 4.4; 4.5; 4.8; new 4.10; 5",
     "Table 12, Figure 9 / statistical_contrasts_bertscore.csv", "RESOLVED"),
    ("R03", "Table / Unsupported claim",
     "Table 9 values for Gemma, Llama and Mistral are all round numbers rising in the same strict order on "
     "every metric while PsyAdapt carries measured three-decimal values; the table looks constructed.",
     "The previous Table 9 was withdrawn and replaced with measured generator-swap results carrying "
     "bootstrap confidence intervals. The text states why the old table was withdrawn. Because the "
     "intervals of all full-context arms overlap, no ordering among generators is claimed.",
     "4.7; Table 9; new Figure 8", "Table 9, Figure 8 / sota_designA.csv", "RESOLVED"),
    ("R04", "Experimental / Consistency",
     "A0 has grounding support 'N/A' in Table 8 because it performs no retrieval, yet the three bare "
     "comparison LLMs in Table 9 each received a grounding score; grounding is defined only against "
     "retrieved passages.",
     "Grounding is now computed only for arms that actually retrieved passages. All bare arms in Table 9 "
     "and A0 in Table 8 show '- (no retrieval)', and the metric definition in 4.2 states that grounding is "
     "undefined without retrieval and that such cells are left blank rather than filled.",
     "4.2 (Grounding Support); Table 8; Table 9", "Table 8, Table 9 / sota_designA.csv, baseline_comparison.csv",
     "RESOLVED"),
    ("R05", "Experimental",
     "The comparison is confounded: Qwen2.5-3B inside the framework is compared against different base "
     "models outside it. The fair test is the framework with each base model swapped in, or each bare "
     "model against bare Qwen.",
     "Both fair tests are now run. Each model is evaluated inside the identical pipeline (same 300 items, "
     "prompts, retrieval cache, decoding settings and metrics) in a bare arm and a full-context arm, and "
     "bare Qwen (A0) is included as the bare-vs-bare reference.",
     "4.7; Table 9", "Table 9 / sota_designA.csv", "RESOLVED"),
    ("R07", "Safety / Ethics / Evaluation",
     "A mental-health support system is evaluated only with BERTScore, ROUGE and embedding cosine "
     "similarity; a blinded human rating on a subset (empathy, helpfulness, safety) is required, ideally "
     "with clinical training, together with an ethics statement and a description of how crisis content is "
     "detected and handled, since 'safety and quality verification' is never described beyond one paragraph.",
     "Three separate additions. (i) The implemented safety mechanisms are now specified: the pre-generation "
     "keyword screen with its three review categories, the safety system prompt, and the post-generation "
     "response screen including the missing-crisis-referral check. (ii) A new subsection states the ethical "
     "position, data provenance, intended use and risks. (iii) The blinded pairwise human-evaluation "
     "protocol is described and its status is reported honestly: no annotations exist, so no human results "
     "are reported or claimed, and the gap is listed as a limitation. The distinction between technical "
     "screening and clinical safety validation is stated in 3.9.2, 3.10, 4.13 and the conclusion.",
     "3.9.2; new 3.10; new 4.13; new 4.14; new 4.16; 5",
     "defect_composition_A7.csv; notebook safety screen, judge and human-eval cells",
     "PARTIALLY RESOLVED - human/clinical evaluation still outstanding"),
    ("R09", "Evaluation / Methodological",
     "No accuracy, F1 or confusion matrix is reported for the emotion, personality or strategy classifiers.",
     "A new subsection reports accuracy, macro-F1, weighted-F1 and macro ROC-AUC for all three components "
     "on their held-out test splits, together with the label-leaking legacy strategy model as an explicit "
     "diagnostic. The text notes that the per-response predictions from which confusion matrices are "
     "derived are exported by the pipeline; the matrices themselves are not reproduced in the paper.",
     "new 4.9; Table 11; 3.2.3; 3.3; 3.7; Table 3; Table 6",
     "Table 11 / component_classifiers.csv", "PARTIALLY RESOLVED - confusion matrices not printed in the paper"),
    ("R10", "Dataset / Citation",
     "Inferring Big Five traits from a single user turn with an SVM is not credible without validation "
     "numbers, and the OCEAN-Chat 'source' link is a Hugging Face search URL, not a dataset.",
     "The validation numbers are now reported (0.52 accuracy, 0.44 macro-F1 for the primary trait) and the "
     "signal is explicitly described as weak, threshold-gated, hedged in the prompt and never shown to the "
     "user. The search URL was removed; because a correct identifier cannot be invented, the source line is "
     "marked UNVERIFIED and flagged for author action.",
     "3.2.3; new 4.9; new 4.16", "Table 11 / component_classifiers.csv",
     "PARTIALLY RESOLVED - dataset citation REQUIRES HUMAN ACTION"),
    ("R11", "Methodological",
     "Need/context detection: no model, no training data, no output format.",
     "The module is now specified as a deterministic rule-based component with no learned parameters: it "
     "scores the normalized message against keyword lexicons and returns primary topic, current need, "
     "urgency level, estimated knowledge level and response preferences, with a documented default when no "
     "lexicon entry matches. The text states that it has not been validated against labelled annotations.",
     "3.6.3; Algorithm 1; new 4.16", "Notebook rule-based context/need/urgency definitions",
     "PARTIALLY RESOLVED - module specified but still unvalidated"),
    ("R12", "Evaluation",
     "'Defects per response' is a core metric but what counts as a defect and how it is detected is never "
     "stated.",
     "The metric definition now describes both implemented procedures: the rule-based screen with its six "
     "pattern categories including missing crisis referral, and the structured annotation protocol run with "
     "a separate larger instruction-tuned model that must quote the offending text, restricted to six "
     "defect categories. A new subsection reports the defect composition for the full configuration and "
     "states that the counts are automated and are not a clinical safety measure.",
     "4.2 (Response Defects); new 4.13; Table 16", "Table 16 / defect_composition_A7.csv", "RESOLVED"),
    ("R13", "Experimental",
     "Memory is a named component and an ablation row, but the benchmark is single-turn so it cannot do "
     "anything. Remove it from the framework and the ablation, or evaluate on multi-turn data.",
     "Memory is retained but evaluated on multi-turn data. A new subsection reports a paired experiment on "
     "40 ESConv test conversations in which the reply to the fourth user turn is generated with and without "
     "memory of the first three turns (0.618 vs 0.584 BERTScore-F1, +0.034, CI [0.001, 0.066], p = 0.048, "
     "d_z = 0.34), described as preliminary. The single-turn row is relabelled a control and the text states "
     "it carries no information about memory.",
     "3.4; Table 6; Table 10 (A5 relabelled); 4.8; new 4.11; Table 13; 5",
     "Table 13 / memory_experiment.csv", "RESOLVED"),
    ("R14", "Dataset",
     "81 chunks is very small for psychoeducational RAG.",
     "The re-chunking procedure is now described (fine passages of about 500 characters with 80-character "
     "overlap at sentence boundaries, built over identical text alongside the original coarse index, "
     "motivated by the ~256-token encoder window), and the small corpus is acknowledged as a limitation. "
     "The chunk counts quoted in Table 2 and Table 6 are marked UNVERIFIED because the pipeline's "
     "chunk-count output was not part of the supplied result set.",
     "3.2.5; Table 2; Table 6; Algorithm 1; new 4.16", "Notebook re-chunking cell",
     "PARTIALLY RESOLVED - corpus still small; chunk counts UNVERIFIED"),
    ("R16", "Evaluation",
     "Grounding support is embedding similarity between response sentences and retrieved passages; a system "
     "that copies retrieved text scores well regardless of whether that text answers the user.",
     "The definition now states that grounding support is an embedding proxy rather than fact-checking, is "
     "computed only against the passages a configuration received, and is undefined without retrieval. The "
     "mean maximum similarity underlying the indicator (0.41) was added to Table 7 so that the indicator "
     "can be interpreted, and 4.12 reports that grounding is the most length-dependent metric.",
     "4.2 (Grounding Support); Table 7; new 4.12; new 4.15; new 4.16",
     "Table 7 / overall_A7.csv; Table 15 / length_metric_correlations.csv", "RESOLVED"),
    ("R17", "Evaluation",
     "BERTScore and ROUGE are measured against MentalChat16K reference responses, and ROUGE-L of 0.105 "
     "means almost no lexical overlap for any configuration, so it discriminates nothing.",
     "Both definitions now state that they measure similarity to a single acceptable reference and not "
     "correctness, helpfulness or safety. The ROUGE-L definition states explicitly that the metric stays "
     "below 0.11 for every configuration, discriminates weakly and is reported for completeness only; the "
     "same point is repeated in the limitations.",
     "4.2 (BERTScore-F1, ROUGE-L F1); new 4.15; new 4.16", "ablation_study.csv, baseline_comparison.csv",
     "RESOLVED"),
    ("R18", "Evaluation / Figure",
     "A7 produces much longer responses than A0; longer outputs inflate BERTScore recall and grounding "
     "coverage, so length should be controlled for or reported alongside.",
     "A new subsection reports the length distribution per configuration and the Spearman correlations "
     "between length and each metric (grounding 0.56, BERTScore-F1 0.41, ROUGE-L 0.29, input relevance "
     "0.18), and states that part of the full configuration's advantage may be a length effect. Figure 7 "
     "and 4.6 were updated to the verified medians. The length-controlled (matched-strata) comparison is "
     "implemented in the pipeline but its output was not supplied, and this is marked PENDING in the text.",
     "4.6; new 4.12; Table 14; Table 15; Figure 7",
     "Table 14 / response_lengths.csv; Table 15 / length_metric_correlations.csv",
     "PARTIALLY RESOLVED - length-controlled contrasts PENDING"),
    ("R20", "Dataset / Reproducibility",
     "MentalChat16K is split 12,784 / 1,596 / 1,598 but nothing appears to be trained on it; state whether "
     "any model is fine-tuned. The 'injection thresholds' taken from the validation split are never "
     "explained.",
     "The text now states that no language model is fine-tuned on MentalChat16K and that Qwen2.5-3B-Instruct "
     "is used zero-shot, and it lists the three purposes the split actually serves. The injection-threshold "
     "procedure is specified as the validation-split confidence quantile that yields a 0.60 injection rate, "
     "with the reason it replaced the earlier fixed 0.35 threshold; a corresponding row was added to Table 6 "
     "and a step to Algorithm 1.",
     "3.2.1; Table 6; Algorithm 1", "Notebook configuration and threshold-selection cells", "RESOLVED"),
    ("R22", "Consistency",
     "A2 is 'RAG + user profile + classification' and A6 is 'full minus strategy guidance'; the difference "
     "between them is not stated, yet they give different results.",
     "The three strategy conditions are now defined explicitly: A2 inserts the classifier-predicted "
     "strategy, A6 omits the strategy block entirely, and A7 replaces the classifier prediction with the "
     "adaptive scorer that combines emotion, intensity, urgency, need and topic. This is stated in the "
     "methodology and repeated in the ablation discussion.",
     "3.7; 4.8", "Notebook configuration definitions", "RESOLVED"),
    ("R23", "Writing",
     "Equation variables are missing throughout the methodology; the equation objects did not survive, so "
     "paragraphs 72 to 147 have blank symbols.",
     "The equation objects are present in the source manuscript, so the blanks arose in conversion. To "
     "remove the risk, every nomenclature sentence and every algorithm step was rewritten with the symbols "
     "as plain text (X, E, P, N, U, S, R, C, Y, Y*, f_emotion, x_i, y_j, s_i, k_j, D_i, T_i, M); the "
     "numbered display equations were left untouched.",
     "3.4; 3.6; 3.7; 3.8; 3.9; 4.2; Algorithm 1", "Original manuscript (OMML objects verified present)",
     "RESOLVED"),
    ("R24", "Writing",
     "Paragraph 11 repeats the abstract almost verbatim, including all five numbers.",
     "The duplicated numeric sentence was removed and replaced with a pointer to Section 4; the numbers now "
     "appear once, in the abstract and in the results.",
     "1. Introduction", "Critical review", "RESOLVED"),
    ("R25", "Writing",
     "There is no Discussion or Limitations section; 'Results and Discussion' is combined and the "
     "Conclusion runs straight after.",
     "A Discussion subsection and a Limitations subsection were added at the end of Section 4, before the "
     "conclusion, so that the existing section numbering is preserved. The discussion states what the "
     "evidence supports; the limitations cover automatic-only evaluation, the single reference, the length "
     "confound, component quality, corpus size, the single-turn benchmark, single-run decoding, statistical "
     "coverage, hardware-specific latency and language scope.",
     "new 4.15; new 4.16", "All supplied result files", "RESOLVED"),
    ("R26", "Writing",
     "Section headings are inconsistent in case ('1.Introduction', '4. RESULTS AND DISCUSSION', "
     "'5. CONCLUSION AND FUTURE WORK').",
     "Headings were normalized to '1. Introduction', '3. Proposed Methodology', '4. Results and Discussion' "
     "and '5. Conclusion and Future Work'. No heading was renamed, reordered or renumbered.",
     "1; 3; 4; 5", "Critical review", "RESOLVED"),
    ("R27", "Citation",
     "Citations are numeric [n]; switch to Author et al. (Year) if the target journal is author-date. All "
     "30 are cited, which is good.",
     "No change made. The requirement is conditional on the target journal's style, which is not specified; "
     "converting 30 references and all in-text citations would be a formatting decision for the authors. "
     "The reference list, numbering and order are unchanged.",
     "6. References (unchanged)", "-", "REQUIRES HUMAN ACTION"),
    ("R29", "Writing",
     "The abstract reads a little mechanically and ends on 'empirical grounds for investigating', which "
     "undersells a finding; once the significance results exist, end on what they show.",
     "The abstract now ends on the finding: the baselines are separated with corrected significance, the "
     "advantage over the retrieval-plus-profiling configuration is not significant, retrieval and profiling "
     "account for most of the gain, emotion and strategy guidance mainly reduce detected defects, and human "
     "evaluation is outstanding. Structure and approximate length were preserved.",
     "Abstract", "statistical_contrasts_bertscore.csv; overall_A7.csv", "RESOLVED"),
    ("R31", "Required action",
     "Rerun Table 9 properly, explain how bare LLMs received a grounding score, and fix the comparison design.",
     "Done under R03, R04 and R05: Table 9 rebuilt from the generator-swap run, grounding restricted to "
     "retrieval arms, and the confound removed.",
     "4.7; Table 9; Figure 8", "sota_designA.csv", "RESOLVED"),
    ("R32", "Required action",
     "Add bootstrap confidence intervals and paired significance tests for every configuration comparison "
     "and rewrite every 'better' claim to match.",
     "Done for BERTScore-F1, the primary metric, in the new statistical subsection, and every comparative "
     "claim was rewritten. Paired tests for the remaining automatic metrics are produced by the same "
     "pipeline procedure but were not part of the supplied result set; this is marked PENDING in the text "
     "rather than filled in.",
     "4.1; new 4.10; Table 12; Figure 9", "statistical_contrasts_bertscore.csv",
     "PARTIALLY RESOLVED - tests for non-BERTScore metrics PENDING"),
    ("R33", "Required action", "Report accuracy for the emotion, personality and strategy classifiers.",
     "Done under R09: Table 11 reports accuracy, macro-F1, weighted-F1 and ROC-AUC for all three, plus the "
     "leaky legacy strategy model as a diagnostic.",
     "new 4.9; Table 11", "component_classifiers.csv", "RESOLVED"),
    ("R34", "Required action",
     "Add a human evaluation on at least 50 to 100 items, plus an ethics and crisis-handling statement.",
     "The ethics and crisis-handling statement was added (new 3.10) and the blinded pairwise protocol is "
     "documented (new 4.14). The human evaluation itself cannot be fabricated: no annotated returns exist, "
     "so the paper reports no human results, claims none, and lists the study as required outstanding work.",
     "new 3.10; new 4.14; new 4.16; 5", "Notebook human-eval export cells (status: pending)",
     "REQUIRES HUMAN ACTION"),
    ("R35", "Required action", "Define the defect detector and the need/context module, or remove them.",
     "Both are now defined: the defect detector in 4.2 and 4.13, the need/context module in 3.6.3. Neither "
     "was removed, and both are described with their limitations.",
     "4.2; 3.6.3; new 4.13", "Notebook screens and judge protocol; defect_composition_A7.csv", "RESOLVED"),
    ("R36", "Required action", "Remove memory from the framework or evaluate it on multi-turn data.",
     "Evaluated on multi-turn data; see R13.", "new 4.11; Table 13", "memory_experiment.csv", "RESOLVED"),
    ("R37", "Required action", "Restore the missing equation symbols and add a Discussion with limitations.",
     "Symbols restored as plain text throughout the methodology, metric definitions and algorithm (R23); "
     "Discussion and Limitations subsections added (R25).",
     "3.4-3.9; 4.2; Algorithm 1; new 4.15; new 4.16", "Original manuscript; all supplied result files",
     "RESOLVED"),
]


def build_matrix():
    d = new_doc(landscape=True)
    heading(d, "PsyAdapt - Reviewer Response Matrix", 15, WD_ALIGN_PARAGRAPH.CENTER)
    body(d, "Manuscript: PsyAdapt: A Personalized LLM-Based Framework for Psychoeducational Dialogue "
            "Generation and Evaluation. Every comment in the critical review receives an explicit "
            "resolution below. Identifiers correspond to the order of the comments in the review document. "
            "Evidence is cited as the paper table or figure together with the result file in tables.zip, or "
            "as the notebook cell that implements the procedure. No result, statistic, participant or "
            "citation was invented: where evidence was unavailable the status says so.")
    table(d,
          ["ID", "Category", "Reviewer comment", "Action taken", "Updated section", "Evidence", "Status"],
          [[m[0], m[1], m[2], m[3], m[4], m[5], m[6]] for m in MATRIX],
          widths=[520, 1000, 3200, 4200, 1500, 1800, 1700], size=8)
    d.add_paragraph()
    heading(d, "Summary of statuses", 12)
    counts = {}
    for m in MATRIX:
        key = m[6].split(" - ")[0]
        counts[key] = counts.get(key, 0) + 1
    body(d, "; ".join(f"{k}: {v}" for k, v in sorted(counts.items())) +
            f" (total comments addressed: {len(MATRIX)}).")
    heading(d, "Items that require author action before submission", 12)
    for txt in [
        "R10 / R27: confirm the OCEAN-Chat dataset identifier and permanent URL, and decide whether the "
        "target journal requires author-date citations.",
        "R14: insert the exact coarse and fine chunk counts from the pipeline's chunk-granularity "
        "statistics output, replacing the UNVERIFIED markers in Table 2, Table 6 and Section 3.2.5.",
        "R18 / R32: run the remaining paired tests and the length-controlled contrasts and replace the "
        "PENDING notes in Sections 4.10 and 4.12.",
        "R34: complete the blinded human annotation (and, if possible, clinician ratings), then add the "
        "results and agreement statistics to Section 4.14.",
        "Figure 6: regenerate the topic-level defect heatmap from the revised per-response output; the "
        "figure currently carries a PENDING note.",
    ]:
        body(d, "\u2022 " + txt)
    d.save(f"{OUT}/REVIEWER_RESPONSE_MATRIX.docx")
    print("wrote REVIEWER_RESPONSE_MATRIX.docx")


# ------------------------------------------------------------------ 2. consistency report
CONSISTENCY = [
    # item, old, new, source file, paper location, status
    ("A7 BERTScore-F1", "0.632", "0.634", "overall_A7.csv / ablation_study.csv",
     "Abstract; Table 7; Table 8; Table 9; Table 10; 4.3-4.5; 4.8; 5", "VERIFIED"),
    ("A7 ROUGE-L F1", "0.105", "0.106", "overall_A7.csv", "Abstract; Tables 7-10; 4.3-4.5; 5", "VERIFIED"),
    ("A7 input relevance", "0.607", "0.609", "overall_A7.csv", "Abstract; Tables 7-10; 4.3-4.5; 5", "VERIFIED"),
    ("A7 grounding support", "0.631", "0.622", "overall_A7.csv", "Abstract; Tables 7-10; 4.3-4.5; 5", "VERIFIED"),
    ("A7 mean-max grounding similarity", "not reported", "0.41", "overall_A7.csv", "Table 7; 4.3", "VERIFIED (new row)"),
    ("A7 defects per response", "0.787", "0.69", "overall_A7.csv / defect_composition_A7.csv",
     "Tables 7-10; 4.3; 4.5; 4.13; 5", "VERIFIED"),
    ("A7 latency (s)", "2.395", "2.3", "overall_A7.csv", "Tables 7-10; 4.3; 4.5; 4.8; 5", "VERIFIED"),
    ("A0 BERTScore-F1 / ROUGE-L / relevance", "0.575 / 0.072 / 0.544", "0.589 / 0.081 / 0.553",
     "baseline_comparison.csv", "Table 8; Table 9; 4.4; 4.7; 4.15", "VERIFIED"),
    ("A0 grounding support", "N/A", "- (no retrieval); metric undefined", "baseline_comparison.csv (blank)",
     "Table 8; Table 9; 4.2; 4.4", "VERIFIED"),
    ("A0 defects / latency", "1.813 / 0.797", "1.41 / 0.81", "baseline_comparison.csv",
     "Table 8; Table 9; 4.4", "VERIFIED"),
    ("A1 BERTScore-F1 / ROUGE-L / relevance", "0.607 / 0.088 / 0.583", "0.613 / 0.095 / 0.581",
     "baseline_comparison.csv", "Table 8; 4.4; 4.15", "VERIFIED"),
    ("A1 grounding / defects / latency", "0.498 / 1.080 / 1.518", "0.482 / 0.98 / 1.44",
     "baseline_comparison.csv", "Table 8; 4.4; 4.15", "VERIFIED"),
    ("A2 full row", "0.631 / 0.103 / 0.604 / 0.630 / 0.789 / 2.296",
     "0.629 / 0.104 / 0.603 / 0.615 / 0.74 / 2.21", "baseline_comparison.csv / ablation_study.csv",
     "Table 8; Table 10; 4.4; 4.8; 4.15", "VERIFIED"),
    ("A3 full row", "0.626 / 0.099 / 0.595 / 0.613 / 0.927 / 2.212",
     "0.621 / 0.099 / 0.592 / 0.598 / 0.86 / 2.12", "ablation_study.csv", "Table 10; 4.8; 4.15", "VERIFIED"),
    ("A4 full row", "0.629 / 0.102 / 0.606 / 0.621 / 0.787 / 2.311",
     "0.626 / 0.102 / 0.601 / 0.608 / 0.76 / 2.19", "ablation_study.csv", "Table 10; 4.8; 4.15", "VERIFIED"),
    ("A5 full row", "0.631 / 0.103 / 0.604 / 0.629 / 0.789 / 2.372",
     "0.629 / 0.104 / 0.603 / 0.614 / 0.74 / 2.20", "ablation_study.csv", "Table 10; 4.8", "VERIFIED"),
    ("A6 full row", "0.625 / 0.099 / 0.599 / 0.629 / 0.907 / 1.950",
     "0.617 / 0.097 / 0.593 / 0.604 / 0.89 / 1.87", "ablation_study.csv", "Table 10; 4.8; 4.15", "VERIFIED"),
    ("A5 row label", "A5 - Without Memory", "A5 - Without Memory (single-turn control)",
     "ablation_study.csv (config_name)", "Table 6; Table 10; Table 14; 3.4; 4.8; 4.11", "VERIFIED"),
    ("Comparison with other LLMs",
     "Gemma 0.610/0.095/0.580/0.590/0.900; Llama 0.615/0.098/0.585/0.600/0.870; "
     "Mistral 0.620/0.100/0.595/0.610/0.850 (round, monotone, grounding for bare models)",
     "Generator-swap arms with bootstrap CIs: full-context Gemma 0.627 [0.621, 0.633], Llama 0.623 "
     "[0.617, 0.629], Mistral 0.630 [0.624, 0.636]; bare arms 0.582 / 0.580 / 0.586 with grounding left "
     "undefined", "sota_designA.csv", "Table 9; Figure 8; 4.7; 5", "VERIFIED (previous values withdrawn)"),
    ("Statistical evidence", "none reported",
     "Paired Wilcoxon, bootstrap CIs, d_z and Holm-corrected p for 7 contrasts; A7-A2 and A7-A5 not "
     "significant", "statistical_contrasts_bertscore.csv", "Table 12; Figure 9; 4.10; Abstract; 5",
     "VERIFIED (new)"),
    ("Component classifier metrics", "none reported",
     "Emotion 0.78/0.72/0.79/0.94; Personality 0.52/0.44/0.51/0.66; Strategy (context-only) "
     "0.61/0.53/0.59/0.72; Strategy (legacy, leaky) 0.87/0.83/0.86/0.95",
     "component_classifiers.csv", "Table 11; 4.9; 3.2.3; 3.7", "VERIFIED (new)"),
    ("Strategy classifier model", "DistilBERT-based classifier",
     "Context-only TF-IDF + logistic regression (validation-selected); legacy leaky variant reported as a "
     "diagnostic only", "component_classifiers.csv; notebook strategy cell",
     "3.3; 3.7; Table 3; Table 6; Table 11; Algorithm 1", "VERIFIED"),
    ("Personality classifier model", "SVM classifier", "TF-IDF + linear SVM classifier",
     "component_classifiers.csv", "Table 3; Table 6; Table 11", "VERIFIED"),
    ("Memory evidence", "single-turn ablation row only",
     "Multi-turn experiment: 0.618 vs 0.584 BERTScore-F1, +0.034, CI [0.001, 0.066], d_z 0.34, p 0.048, "
     "40 conversations", "memory_experiment.csv", "Table 13; 4.11; 3.4; 5", "VERIFIED (new)"),
    ("Response length medians (A0 / A1 / A7)", "159 / 166 / 179 words", "128 / 141 / 172 words",
     "response_lengths.csv", "4.6; Table 14; Figure 7", "VERIFIED"),
    ("Length-metric correlations", "not reported",
     "BERTScore-F1 0.41; ROUGE-L 0.29; input relevance 0.18; grounding 0.56", "length_metric_correlations.csv",
     "Table 15; 4.6; 4.12", "VERIFIED (new)"),
    ("Defect composition (A7)", "not reported",
     "0.21 / 0.14 / 0.12 / 0.09 / 0.08 / 0.05, total 0.69", "defect_composition_A7.csv",
     "Table 16; 4.13", "VERIFIED (new)"),
    ("Injection thresholds", "fixed 0.35, unexplained",
     "Validation-split confidence quantile targeting a 0.60 injection rate", "Notebook threshold cell",
     "3.2.1; Table 6; Algorithm 1", "VERIFIED against code (not a numeric result)"),
    ("Knowledge-base chunking", "81 indexed text chunks",
     "81 coarse chunks re-chunked into ~500-character fine passages (80-character overlap); fine index used "
     "for all reported configurations", "Notebook re-chunking cell",
     "3.2.5; Table 2; Table 6; Algorithm 1", "UNVERIFIED chunk counts - author must insert the pipeline value"),
    ("Figure 4 / Figure 5 / Figure 7 images", "plotted from the previous run",
     "regenerated from the supplied result files (same figure numbers, same layout)",
     "overall_A7.csv; ablation_study.csv; baseline_comparison.csv; response_lengths.csv",
     "Figures 4, 5, 7", "VERIFIED"),
    ("Figure 6 (defects by topic)", "plotted from the previous run", "unchanged; PENDING note added to the caption",
     "not present in tables.zip", "Figure 6", "NOT VERIFIED - regeneration required"),
    ("Human-evaluation results", "none", "none reported; protocol documented and status stated",
     "notebook human-eval status (no annotator files)", "4.14; 4.16; 5",
     "PENDING - REQUIRES HUMAN EVALUATION"),
    ("Paired tests for metrics other than BERTScore-F1", "none", "not reported; PENDING note in the text",
     "not present in tables.zip", "4.10", "PENDING"),
    ("Length-controlled contrasts", "none", "not reported; PENDING note in the text",
     "not present in tables.zip", "4.12", "PENDING"),
    ("Reference list", "30 references, numeric style", "unchanged (numbering, order and entries preserved)",
     "-", "6. References", "VERIFIED unchanged"),
]


def build_consistency():
    d = new_doc(landscape=True)
    heading(d, "PsyAdapt - Final Consistency Report", 15, WD_ALIGN_PARAGRAPH.CENTER)
    body(d, "Audit of the revised manuscript against the supplied updated result files (tables.zip) and the "
            "reviewer-revision notebook. Every value that changed is listed with its source file and the "
            "locations in the paper where it appears, so that no stale value survives in the abstract, "
            "results, discussion, conclusion, captions or tables. Items that could not be verified from the "
            "supplied evidence are listed as UNVERIFIED or PENDING rather than being estimated.")
    table(d, ["Item", "Old value", "New value", "Source file", "Paper location", "Verification status"],
          [list(r) for r in CONSISTENCY],
          widths=[1900, 2400, 3500, 2000, 2500, 2140], size=8)
    d.add_paragraph()
    heading(d, "Cross-document checks performed", 12)
    for txt in [
        "Every A0-A7 metric in Tables 7, 8, 9, 10 and 14 was compared cell by cell with "
        "baseline_comparison.csv, ablation_study.csv, overall_A7.csv, sota_designA.csv and "
        "response_lengths.csv; all values match.",
        "The abstract, introduction, Sections 4.3 to 4.8, the new Sections 4.9 to 4.16 and the conclusion "
        "were searched for the superseded values 0.632, 0.105, 0.607, 0.631, 0.787, 2.395, 1.813, 1.080, "
        "0.498, 0.575, 0.072, 0.544, 0.927, 0.907, 2.296, 2.372, 2.212, 2.311, 1.950 and for the withdrawn "
        "comparison values 0.610, 0.615, 0.620, 0.900, 0.870, 0.850: none remain.",
        "Statistical statements were checked against statistical_contrasts_bertscore.csv: no contrast whose "
        "Holm-corrected p exceeds 0.05 is described as significant or as an improvement anywhere in the "
        "manuscript.",
        "Grounding support appears only for configurations and arms that performed retrieval; A0 and all "
        "bare arms show '- (no retrieval)'.",
        "Configuration names, the count of benchmark items (300), the dataset sizes (15,978 rows; 15,856 "
        "unique inputs; 12,784 / 1,596 / 1,598 split) and the model identifiers were left unchanged and "
        "are consistent across Tables 2, 3, 6 and the text.",
        "Figure and table numbering is unchanged for Figures 1-7 and Tables 1-10; new material uses "
        "Figures 8-9 and Tables 11-16 so that no existing number shifts.",
        "The equation objects in the source manuscript were verified to be present (66 OMML objects); the "
        "nomenclature sentences were additionally written in plain text so the symbols survive conversion.",
    ]:
        body(d, "\u2022 " + txt)
    d.save(f"{OUT}/FINAL_CONSISTENCY_REPORT.docx")
    print("wrote FINAL_CONSISTENCY_REPORT.docx")


# ------------------------------------------------------------------ 3. updated tables
def build_tables():
    out = f"{OUT}/UPDATED_TABLES"
    os.makedirs(out, exist_ok=True)
    payloads = {
        "table07_overall_performance_A7.csv": (["Evaluation Metric", "A7 - PsyAdapt"],
                                               [["BERTScore-F1", "0.634"], ["ROUGE-L F1", "0.106"],
                                                ["Input Relevance", "0.609"],
                                                ["Grounding Support (tau = 0.50)", "0.622"],
                                                ["Mean-max grounding similarity", "0.41"],
                                                ["Defects per Response", "0.69"], ["Latency (s)", "2.3"]]),
        "table08_baseline_comparison.csv": (
            ["Configuration", "BERTScore-F1", "ROUGE-L F1", "Input Relevance", "Grounding Support",
             "Defects/Response", "Latency (s)"],
            [["A0 - Generic LLM", "0.589", "0.081", "0.553", "- (no retrieval)", "1.41", "0.81"],
             ["A1 - LLM + Standard RAG", "0.613", "0.095", "0.581", "0.482", "0.98", "1.44"],
             ["A2 - RAG + User Profile + Classifier Strategy", "0.629", "0.104", "0.603", "0.615", "0.74", "2.21"],
             ["A7 - Full PsyAdapt", "0.634", "0.106", "0.609", "0.622", "0.69", "2.30"]]),
        "table09_sota_generator_swap.csv": (R.T9_HEADER, R.T9_ROWS),
        "table10_ablation_study.csv": (
            ["Configuration", "BERTScore-F1", "ROUGE-L F1", "Input Relevance", "Grounding Support",
             "Defects/Response", "Latency (s)"],
            [["A2 - RAG + User Profile + Classifier Strategy", "0.629", "0.104", "0.603", "0.615", "0.74", "2.21"],
             ["A3 - Without Emotion Detection", "0.621", "0.099", "0.592", "0.598", "0.86", "2.12"],
             ["A4 - Without Personality Detection", "0.626", "0.102", "0.601", "0.608", "0.76", "2.19"],
             ["A5 - Without Memory (single-turn control)", "0.629", "0.104", "0.603", "0.614", "0.74", "2.20"],
             ["A6 - Without Strategy Guidance", "0.617", "0.097", "0.593", "0.604", "0.89", "1.87"],
             ["A7 - Full PsyAdapt", "0.634", "0.106", "0.609", "0.622", "0.69", "2.30"]]),
        "table11_component_classifiers.csv": (R.T11_HEADER, R.T11_ROWS),
        "table12_paired_contrasts_bertscore.csv": (R.T12_HEADER, R.T12_ROWS),
        "table13_memory_experiment.csv": (R.T13_HEADER, R.T13_ROWS),
        "table14_response_lengths.csv": (R.T14_HEADER, R.T14_ROWS),
        "table15_length_metric_correlations.csv": (R.T15_HEADER, R.T15_ROWS),
        "table16_defect_composition_A7.csv": (R.T16_HEADER, R.T16_ROWS),
    }
    for name, (header, rows) in payloads.items():
        with open(os.path.join(out, name), "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(header)
            w.writerows(rows)
    srcdir = os.path.join(out, "source_result_files_from_tables_zip")
    os.makedirs(srcdir, exist_ok=True)
    for f in sorted(os.listdir(TSRC)):
        if f.endswith(".csv"):
            shutil.copy2(os.path.join(TSRC, f), os.path.join(srcdir, f))
    print("wrote UPDATED_TABLES/ with", len(payloads), "paper tables +",
          len(os.listdir(srcdir)), "source files")


def build_figure_readme():
    src = "/projects/sandbox/GCS/_work/orig_fig/word/media"
    dst = f"{OUT}/UPDATED_FIGURES"
    for n, name in [(1, "figure1_framework_architecture.png"), (2, "figure2_user_representation.png"),
                    (3, "figure3_overall_architecture.png"), (6, "figure6_defects_by_topic_PENDING.png")]:
        shutil.copy2(os.path.join(src, f"image{n}.png"), os.path.join(dst, name))
    with open(os.path.join(dst, "README.txt"), "w", encoding="utf-8") as f:
        f.write(
            "Final figures used in the revised manuscript\n"
            "===========================================\n\n"
            "figure1_framework_architecture.png   Figure 1 - unchanged (schematic, no measured values)\n"
            "figure2_user_representation.png      Figure 2 - unchanged (schematic)\n"
            "figure3_overall_architecture.png     Figure 3 - unchanged (schematic)\n"
            "figure4_overall_performance_A7.png   Figure 4 - REGENERATED from overall_A7.csv\n"
            "figure5_quality_vs_latency.png       Figure 5 - REGENERATED from baseline_comparison.csv + "
            "ablation_study.csv\n"
            "figure6_defects_by_topic_PENDING.png Figure 6 - NOT regenerated: the topic-level defect data "
            "were not supplied in tables.zip. The figure in the manuscript still shows the previous run and "
            "carries a PENDING note in its caption. Replace it with the export from the revision notebook.\n"
            "figure7_response_lengths.png         Figure 7 - REGENERATED from response_lengths.csv "
            "(median, interquartile range and mean; whiskers were not drawn because only the quartiles were "
            "supplied, and the caption/axis wording reflects this)\n"
            "figure8_sota_design_a.png            Figure 8 - NEW, from sota_designA.csv (means with 95% "
            "bootstrap confidence intervals)\n"
            "figure9_paired_contrasts.png         Figure 9 - NEW, from statistical_contrasts_bertscore.csv "
            "(paired mean differences with 95% bootstrap confidence intervals)\n\n"
            "All regenerated panels keep the aspect ratio of the figure they replace so that the layout of "
            "the manuscript is unchanged.\n")
    print("wrote UPDATED_FIGURES/README.txt and copied unchanged figures")


if __name__ == "__main__":
    build_matrix()
    build_consistency()
    build_tables()
    build_figure_readme()
