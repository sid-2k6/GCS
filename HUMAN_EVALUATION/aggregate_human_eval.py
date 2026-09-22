"""
Score returned human annotations and emit the paper-ready table for Section 4.14.

Reads   : human_pairwise_key_DO_NOT_SHARE.csv  +  every human_pairwise_<initials>.csv in --dir
Writes  : human_pairwise_winrates.csv          per contrast x criterion, with exact binomial p and Holm
          human_agreement.csv                  pairwise raw agreement and Cohen's kappa per annotator pair
          table17_human_evaluation.csv         paper-ready (drop into UPDATED_TABLES/)
          human_vs_judge_agreement.csv         optional, if pairwise_judgments.jsonl is given

Nothing is imputed. Blank cells are ignored, and any contrast x criterion with fewer than 8 non-tied
judgments is reported with p = NA rather than a fabricated test result.

Usage
-----
    python aggregate_human_eval.py --dir . [--judge /path/to/pairwise_judgments.jsonl]
"""
import argparse
import glob
import itertools
import json
import os

import numpy as np
import pandas as pd
from scipy.stats import binomtest

CRITERIA = ["relevance", "empathy", "personalization", "helpfulness", "safety", "overall"]
MIN_N_FOR_TEST = 8


def holm(pvals):
    p = np.asarray(pvals, dtype=float)
    out = np.full_like(p, np.nan)
    idx = np.where(~np.isnan(p))[0]
    if len(idx) == 0:
        return out
    order = idx[np.argsort(p[idx])]
    m = len(order)
    running = 0.0
    for rank, i in enumerate(order):
        val = (m - rank) * p[i]
        running = max(running, min(val, 1.0))
        out[i] = running
    return out


def cohens_kappa(a, b):
    labels = sorted(set(a) | set(b))
    n = len(a)
    if n == 0 or len(labels) < 2:
        return np.nan
    po = np.mean([x == y for x, y in zip(a, b)])
    pe = sum((np.mean([x == l for x in a]) * np.mean([y == l for y in b])) for l in labels)
    return np.nan if pe >= 1 else (po - pe) / (1 - pe)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", default=".")
    ap.add_argument("--judge", default=None, help="evaluation/pairwise_judgments.jsonl (optional)")
    a = ap.parse_args()

    key_path = os.path.join(a.dir, "human_pairwise_key_DO_NOT_SHARE.csv")
    if not os.path.exists(key_path):
        raise SystemExit("key file not found - run build_human_eval_sheet.py first")
    key = pd.read_csv(key_path)

    files = [f for f in glob.glob(os.path.join(a.dir, "human_pairwise_*.csv"))
             if "key" not in os.path.basename(f) and "BLINDED" not in os.path.basename(f)]
    if not files:
        raise SystemExit("no annotator files found (expected human_pairwise_<initials>.csv)")

    frames = []
    for f in files:
        df = pd.read_csv(f)
        df["annotator"] = os.path.basename(f).replace("human_pairwise_", "").replace(".csv", "")
        frames.append(df)
    ann = pd.concat(frames, ignore_index=True).merge(key, on="item_id", how="inner")
    print(f"{len(files)} annotator file(s), {ann['item_id'].nunique()} items, {len(ann)} judgements")

    # map A/B choices onto configuration names
    for c in CRITERIA:
        if c not in ann:
            ann[c] = np.nan
        v = ann[c].astype(str).str.strip().str.lower()
        ann[f"winner_{c}"] = np.where(
            v.eq("tie"), "tie",
            np.where(v.eq("a"), ann["first_config"],
                     np.where(v.eq("b"), ann["second_config"], np.nan)))

    # ---------------- win rates ----------------
    rows = []
    for contrast, grp in ann.groupby("contrast"):
        x, y = contrast.split("_vs_")
        for c in CRITERIA:
            w = grp[f"winner_{c}"].dropna()
            nx, ny = int((w == x).sum()), int((w == y).sum())
            ties = int((w == "tie").sum())
            decided = nx + ny
            p = binomtest(nx, decided, 0.5).pvalue if decided >= MIN_N_FOR_TEST else np.nan
            rows.append({"contrast": contrast, "criterion": c, "config_a": x, "config_b": y,
                         "n_judgements": len(w), "wins_a": nx, "wins_b": ny, "ties": ties,
                         "win_rate_a_of_decided": round(nx / decided, 4) if decided else np.nan,
                         "p_value_exact_binomial": p})
    wr = pd.DataFrame(rows)
    wr["p_holm_within_criterion"] = np.nan
    for c, g in wr.groupby("criterion"):
        wr.loc[g.index, "p_holm_within_criterion"] = holm(g["p_value_exact_binomial"].values)
    wr.to_csv(os.path.join(a.dir, "human_pairwise_winrates.csv"), index=False)

    # ---------------- agreement ----------------
    agr = []
    if ann["annotator"].nunique() >= 2:
        for c in CRITERIA:
            piv = ann.pivot_table(index="item_id", columns="annotator",
                                  values=f"winner_{c}", aggfunc="first")
            for p1, p2 in itertools.combinations(piv.columns, 2):
                both = piv[[p1, p2]].dropna()
                if len(both) == 0:
                    continue
                agr.append({"criterion": c, "annotator_a": p1, "annotator_b": p2, "n_items": len(both),
                            "raw_agreement": round(float((both[p1] == both[p2]).mean()), 4),
                            "cohens_kappa": round(float(cohens_kappa(both[p1].tolist(), both[p2].tolist())), 4)})
        pd.DataFrame(agr).to_csv(os.path.join(a.dir, "human_agreement.csv"), index=False)
    else:
        print("! only one annotator: no agreement statistics, and none may be reported in the paper")

    # ---------------- paper-ready table ----------------
    paper = wr.copy()
    paper["Contrast"] = paper["config_a"] + " vs " + paper["config_b"]
    paper["Wins (first)"] = paper["wins_a"]
    paper["Wins (second)"] = paper["wins_b"]
    paper["Win rate"] = paper["win_rate_a_of_decided"].map(
        lambda v: "-" if pd.isna(v) else f"{v:.2f}")
    paper["Exact binomial p"] = paper["p_value_exact_binomial"].map(
        lambda v: "-" if pd.isna(v) else f"{v:.3f}")
    paper["Holm-corrected p"] = paper["p_holm_within_criterion"].map(
        lambda v: "-" if pd.isna(v) else f"{v:.3f}")
    paper["Reading"] = np.where(paper["p_holm_within_criterion"] < 0.05, "Significant",
                                np.where(paper["p_holm_within_criterion"].isna(), "Not tested (n too small)",
                                         "Not significant"))
    paper = paper[["Contrast", "criterion", "n_judgements", "Wins (first)", "Wins (second)", "ties",
                   "Win rate", "Exact binomial p", "Holm-corrected p", "Reading"]]
    paper.columns = ["Contrast", "Criterion", "Judgements", "Wins (first config)", "Wins (second config)",
                     "Ties", "Win rate (of decided)", "Exact binomial p", "Holm-corrected p", "Reading"]
    paper.to_csv(os.path.join(a.dir, "table17_human_evaluation.csv"), index=False)

    # ---------------- human vs LLM judge ----------------
    if a.judge and os.path.exists(a.judge):
        recs = [json.loads(l) for l in open(a.judge) if l.strip()]
        j = pd.DataFrame(recs)
        if {"first_config", "second_config", "overall", "sample_id", "contrast"} <= set(j.columns):
            j["winner"] = np.where(j["overall"].eq("tie"), "tie",
                                   np.where(j["overall"].eq("a"), j["first_config"], j["second_config"]))
            jm = j.groupby(["contrast", "sample_id"])["winner"].agg(lambda s: s.mode().iat[0])
            ann["sample_id"] = ann["sample_id"].astype(str)
            hm = ann.dropna(subset=["winner_overall"]).groupby(["contrast", "sample_id"])["winner_overall"] \
                    .agg(lambda s: s.mode().iat[0])
            both = pd.concat([hm.rename("human"), jm.rename("judge")], axis=1).dropna()
            out = pd.DataFrame([{"n_items": len(both),
                                 "human_judge_raw_agreement": round(float((both.human == both.judge).mean()), 4),
                                 "cohens_kappa": round(float(cohens_kappa(both.human.tolist(),
                                                                          both.judge.tolist())), 4)}])
            out.to_csv(os.path.join(a.dir, "human_vs_judge_agreement.csv"), index=False)
            print("human vs judge agreement:", out.to_dict("records")[0])

    print("\nwrote human_pairwise_winrates.csv, human_agreement.csv, table17_human_evaluation.csv")
    print("Report only what these files contain. If a criterion shows 'Not tested', say so in the paper.")


if __name__ == "__main__":
    main()
