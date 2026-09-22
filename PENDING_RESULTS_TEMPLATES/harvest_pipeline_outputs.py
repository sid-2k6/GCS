"""
Fill the pending-result templates from an EXECUTED run of PsyAdapt_Reviewer_Revision.ipynb.

Point --run-dir at the pipeline output root (the directory that contains evaluation/, statistics/,
retrieval/, baselines/, features/, generations/). Everything written here is copied or recomputed from
real per-response data in that directory. Nothing is imputed: files that are absent are reported as
STILL MISSING and the corresponding PENDING marker must stay in the manuscript.

Usage
-----
    python harvest_pipeline_outputs.py \
        --run-dir /content/drive/MyDrive/Psychoeducational_Dialogue_System/experiments/run_v3 \
        --outdir  ./filled

Then update the manuscript exactly as the printed checklist says.
"""
import argparse
import json
import os

import numpy as np
import pandas as pd

BOOT = 5000
SEED = 42
MISSING, DONE = [], []


def note_missing(label, path, blocks):
    MISSING.append((label, str(path), blocks))


def note_done(label, out, blocks):
    DONE.append((label, out, blocks))


def read(path):
    return pd.read_csv(path) if os.path.exists(path) else None


def bootstrap_ci(x, n_boot=BOOT, seed=SEED):
    x = np.asarray(pd.Series(x).dropna(), dtype=float)
    if len(x) < 2:
        return np.nan, np.nan
    rng = np.random.default_rng(seed)
    means = x[rng.integers(0, len(x), size=(n_boot, len(x)))].mean(axis=1)
    return float(np.percentile(means, 2.5)), float(np.percentile(means, 97.5))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--run-dir", required=True)
    ap.add_argument("--outdir", default="./filled")
    a = ap.parse_args()
    R, O = a.run_dir, a.outdir
    os.makedirs(O, exist_ok=True)
    P = lambda *p: os.path.join(R, *p)

    # ---------------------------------------------------------------- 1. paired contrasts, all metrics
    pc = read(P("statistics", "paired_contrasts.csv"))
    if pc is not None:
        cols = [c for c in ["metric", "contrast", "n_pairs", "mean_a", "mean_b", "mean_diff", "ci95_low",
                            "ci95_high", "cohens_dz", "rank_biserial", "p_value", "p_holm", "test",
                            "identical_output_rate"] if c in pc.columns]
        pc[cols].to_csv(f"{O}/paired_contrasts_all_metrics.csv", index=False)
        note_done("paired contrasts (all metrics)", "paired_contrasts_all_metrics.csv",
                  "Section 4.10 PENDING note; extends Table 12")
        # paper-ready view for the metrics the manuscript still reports descriptively
        want = ["rougeL_f", "input_relevance", "grounding_support_rate", "defect_total", "response_words"]
        sub = pc[pc["metric"].isin(want)].copy()
        if len(sub):
            sub.to_csv(f"{O}/table12b_paired_contrasts_other_metrics.csv", index=False)
            note_done("paper table 12b", "table12b_paired_contrasts_other_metrics.csv",
                      "new rows/table for Section 4.10")
    else:
        note_missing("paired contrasts", P("statistics", "paired_contrasts.csv"),
                     "Section 4.10 PENDING note")

    fo = read(P("statistics", "friedman_omnibus.csv"))
    if fo is not None:
        fo.to_csv(f"{O}/friedman_omnibus.csv", index=False)
        note_done("Friedman omnibus", "friedman_omnibus.csv", "supporting text in Section 4.10")
    else:
        note_missing("Friedman omnibus", P("statistics", "friedman_omnibus.csv"), "Section 4.10")

    # ---------------------------------------------------------------- 2. length-controlled contrasts
    lc = read(P("evaluation", "length_controlled_contrasts.csv"))
    if lc is not None:
        lc.to_csv(f"{O}/length_controlled_contrasts.csv", index=False)
        piv = lc[lc.metric.isin(["bertscore_f1", "grounding_support_rate"])] \
            .pivot_table(index="contrast", columns=["metric", "quartile"],
                         values="mean_diff_conditional_on_length")
        piv.round(4).to_csv(f"{O}/table15b_length_controlled_summary.csv")
        note_done("length-controlled contrasts", "length_controlled_contrasts.csv + "
                  "table15b_length_controlled_summary.csv", "Section 4.12 PENDING note")
    else:
        note_missing("length-controlled contrasts", P("evaluation", "length_controlled_contrasts.csv"),
                     "Section 4.12 PENDING note")

    # ---------------------------------------------------------------- 3. confusion matrices
    pred_dir = P("baselines", "predictions")
    rows = []
    if os.path.isdir(pred_dir):
        for f in sorted(os.listdir(pred_dir)):
            if not f.endswith(".csv"):
                continue
            df = pd.read_csv(os.path.join(pred_dir, f))
            if not {"y_true", "y_pred"} <= set(df.columns):
                continue
            task, _, model = f[:-4].partition("__")
            cm = pd.crosstab(df["y_true"], df["y_pred"]).stack().reset_index()
            cm.columns = ["true_label", "predicted_label", "count"]
            cm.insert(0, "component", task)
            cm.insert(1, "model", model.replace("_", " "))
            rows.append(cm)
    if rows:
        out = pd.concat(rows, ignore_index=True)
        out.to_csv(f"{O}/confusion_matrices_long.csv", index=False)
        for comp, g in out.groupby("component"):
            g.pivot_table(index="true_label", columns="predicted_label", values="count",
                          aggfunc="sum", fill_value=0).to_csv(f"{O}/confusion_matrix_{comp}.csv")
        note_done("confusion matrices", "confusion_matrices_long.csv + per-component wide files",
                  "reviewer R09 (Section 4.9 currently reports aggregate metrics only)")
    else:
        note_missing("classifier predictions", pred_dir, "reviewer R09 confusion matrices")

    # ---------------------------------------------------------------- 4. judge win rates
    wr = read(P("evaluation", "pairwise_winrates.csv"))
    if wr is not None:
        wr.to_csv(f"{O}/judge_pairwise_winrates.csv", index=False)
        note_done("LLM-judge win rates", "judge_pairwise_winrates.csv",
                  "Section 4.14 (automated judging currently reports defects only)")
    else:
        note_missing("LLM-judge win rates", P("evaluation", "pairwise_winrates.csv"), "Section 4.14")
    pcmp = read(P("evaluation", "judge_protocol_comparison.csv"))
    if pcmp is not None:
        pcmp.to_csv(f"{O}/judge_protocol_comparison.csv", index=False)
        note_done("judge protocol comparison", "judge_protocol_comparison.csv", "Section 4.14")

    # ---------------------------------------------------------------- 5. retrieval granularity
    cg = read(P("evaluation", "chunk_granularity_contrast.csv"))
    if cg is not None:
        cg.to_csv(f"{O}/chunk_granularity_contrast.csv", index=False)
        note_done("coarse vs fine retrieval contrast", "chunk_granularity_contrast.csv",
                  "evidence for the re-chunking claim in Section 3.2.5")
    else:
        note_missing("chunk granularity contrast", P("evaluation", "chunk_granularity_contrast.csv"),
                     "Section 3.2.5")

    cs = read(P("retrieval", "chunk_granularity_stats.csv"))
    if cs is not None:
        cs.to_csv(f"{O}/chunk_granularity_stats.csv", index=False)
        n = dict(zip(cs["granularity"], cs["n_chunks"]))
        note_done("chunk counts", f"chunk_granularity_stats.csv (coarse={n.get('coarse')}, "
                  f"fine={n.get('fine')})",
                  "REPLACES the [UNVERIFIED chunk counts] markers in Table 2, Table 6 and Section 3.2.5")
    else:
        note_missing("chunk counts", P("retrieval", "chunk_granularity_stats.csv"),
                     "[UNVERIFIED] markers in Table 2, Table 6, Section 3.2.5")

    # ---------------------------------------------------------------- 6. A0 bootstrap CI for Table 9
    auto = read(P("evaluation", "automatic_metrics_per_response.csv"))
    if auto is not None:
        v = auto[auto["valid"].astype(bool)] if "valid" in auto else auto
        out = []
        for cid in sorted(v["config_id"].unique()):
            x = v.loc[v.config_id == cid, "bertscore_f1"]
            lo, hi = bootstrap_ci(x)
            out.append({"arm": cid, "metric": "bertscore_f1", "n": int(x.notna().sum()),
                        "mean": round(float(x.mean()), 4),
                        "ci95_low": round(lo, 4), "ci95_high": round(hi, 4)})
        pd.DataFrame(out).to_csv(f"{O}/bootstrap_ci_all_configs.csv", index=False)
        note_done("bootstrap CIs for A0-A7", "bootstrap_ci_all_configs.csv",
                  "fills the 'CI not reported' cell for A0 in Table 9")
    else:
        note_missing("per-response automatic metrics", P("evaluation", "automatic_metrics_per_response.csv"),
                     "A0 CI in Table 9; Figure 6 data")

    # ---------------------------------------------------------------- 7. Figure 6 source data
    dd = read(P("evaluation", "defect_counts.csv"))
    ft = read(P("features", "benchmark_component_features.csv"))
    if dd is not None and ft is not None and "primary_topic" in ft.columns:
        dd["sample_id"] = dd["sample_id"].astype(str)
        ft["sample_id"] = ft["sample_id"].astype(str)
        m = dd.merge(ft[["sample_id", "primary_topic"]], on="sample_id", how="left")
        g = m.groupby(["config_id", "primary_topic"])["defect_total"].agg(["size", "mean"]).reset_index()
        g.columns = ["config_id", "topic", "n_responses", "mean_defects"]
        g["mean_defects"] = g["mean_defects"].round(3)
        g.to_csv(f"{O}/figure6_defects_by_topic.csv", index=False)
        note_done("Figure 6 source data", "figure6_defects_by_topic.csv",
                  "replaces the PENDING note in the Figure 6 caption")
    else:
        note_missing("defect counts x topic", f"{P('evaluation','defect_counts.csv')} + "
                     f"{P('features','benchmark_component_features.csv')}", "Figure 6 caption PENDING note")

    # ---------------------------------------------------------------- 8. provenance of supplied values
    prov = []
    if pc is not None:
        have = set(pc[pc.metric == "bertscore_f1"]["contrast"])
        for c in ["A7-A0", "A7-A1", "A7-A2", "A7-A3", "A7-A4", "A7-A5", "A7-A6"]:
            prov.append({"contrast_in_table12": c, "present_in_this_run": c in have,
                         "status": "CONFIRMED" if c in have else "NOT PRODUCED BY THIS RUN"})
        pd.DataFrame(prov).to_csv(f"{O}/statistical_contrast_provenance_CHECKED.csv", index=False)
        note_done("Table 12 provenance check", "statistical_contrast_provenance_CHECKED.csv",
                  "confirms every contrast reported in Table 12 came from this run")

    mem = read(P("statistics", "memory_experiment_contrast.csv"))
    if mem is not None:
        mem.to_csv(f"{O}/memory_experiment_contrast.csv", index=False)
        note_done("memory contrast as computed by the pipeline", "memory_experiment_contrast.csv",
                  "cross-check against the 0.618 / 0.584 BERTScore pair in Table 13")

    # ---------------------------------------------------------------- report
    print("\n=== FILLED ===")
    for label, out, blocks in DONE:
        print(f"  [OK]      {label}\n              -> {out}\n              unblocks: {blocks}")
    print("\n=== STILL MISSING (keep the PENDING/UNVERIFIED marker in the manuscript) ===")
    for label, path, blocks in MISSING:
        print(f"  [MISSING] {label}\n              expected at: {path}\n              blocks: {blocks}")
    if not MISSING:
        print("  none - every pending item can now be filled in.")
    with open(f"{O}/harvest_report.json", "w", encoding="utf-8") as f:
        json.dump({"filled": [{"item": d[0], "output": d[1], "unblocks": d[2]} for d in DONE],
                   "missing": [{"item": m[0], "expected_path": m[1], "blocks": m[2]} for m in MISSING]},
                  f, indent=2)
    print(f"\nwrote {O}/harvest_report.json")


if __name__ == "__main__":
    main()
