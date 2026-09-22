"""
Build the blinded pairwise human-evaluation sheet for PsyAdapt.

Run this INSIDE the executed pipeline session (or anywhere the two input files are readable).
It reads real generated responses and writes:

    human_pairwise_sheet_BLINDED.csv        -> give this to annotators (no system identities)
    human_pairwise_key_DO_NOT_SHARE.csv     -> keep this private; needed to score the returns

No values are invented: every row is copied from the generation table.

Inputs
------
GEN_FINAL : generations/generations_final.csv   (columns: sample_id, config_id, response, valid)
BENCHMARK : the fixed 300-item benchmark CSV    (columns: sample_id, user_input)

Usage
-----
    python build_human_eval_sheet.py \
        --generations /content/drive/MyDrive/.../run_v3/generations/generations_final.csv \
        --benchmark   /content/drive/MyDrive/.../benchmark_300.csv \
        --n-items 100 --seed 43 --outdir .
"""
import argparse
import os

import numpy as np
import pandas as pd

CRITERIA = ["relevance", "empathy", "personalization", "helpfulness", "safety", "overall"]
# Same three contrasts the manuscript reports as exported for annotation (Section 4.14).
CONTRASTS = [("A7", "A2"), ("A2", "A1"), ("A2", "A6")]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--generations", required=True)
    ap.add_argument("--benchmark", required=True)
    ap.add_argument("--n-items", type=int, default=100,
                    help="total items across all contrasts (reviewer asked for 50-100)")
    ap.add_argument("--seed", type=int, default=43)
    ap.add_argument("--outdir", default=".")
    ap.add_argument("--n-annotators", type=int, default=3,
                    help="also writes one identical copy per annotator for convenience")
    a = ap.parse_args()

    gen = pd.read_csv(a.generations)
    gen["sample_id"] = gen["sample_id"].astype(str)
    if "valid" in gen:
        gen = gen[gen["valid"].astype(bool)]
    resp = gen.set_index(["sample_id", "config_id"])["response"]

    bench = pd.read_csv(a.benchmark)
    bench["sample_id"] = bench["sample_id"].astype(str)
    user = bench.set_index("sample_id")["user_input"]

    rng = np.random.default_rng(a.seed)
    per_contrast = int(np.ceil(a.n_items / len(CONTRASTS)))
    rows = []
    for x, y in CONTRASTS:
        # only samples where both configurations produced different, non-empty text carry information
        usable = []
        for sid in user.index:
            if (sid, x) not in resp.index or (sid, y) not in resp.index:
                continue
            rx, ry = str(resp.loc[(sid, x)]).strip(), str(resp.loc[(sid, y)]).strip()
            if rx and ry and rx != ry:
                usable.append(sid)
        if len(usable) < per_contrast:
            print(f"! {x}_vs_{y}: only {len(usable)} usable items (asked for {per_contrast})")
        take = rng.choice(usable, size=min(per_contrast, len(usable)), replace=False)
        for sid in sorted(take):
            flip = bool(rng.integers(0, 2))
            first, second = (y, x) if flip else (x, y)
            rows.append({"sample_id": sid, "contrast": f"{x}_vs_{y}",
                         "first_config": first, "second_config": second,
                         "user_message": user.loc[sid],
                         "reply_A": resp.loc[(sid, first)], "reply_B": resp.loc[(sid, second)]})

    df = pd.DataFrame(rows).sample(frac=1.0, random_state=a.seed).reset_index(drop=True)
    df = df.iloc[: a.n_items].copy()
    df["item_id"] = [f"H{i:04d}" for i in range(1, len(df) + 1)]

    os.makedirs(a.outdir, exist_ok=True)
    key = df[["item_id", "sample_id", "contrast", "first_config", "second_config"]]
    key.to_csv(os.path.join(a.outdir, "human_pairwise_key_DO_NOT_SHARE.csv"), index=False)

    sheet = df[["item_id", "user_message", "reply_A", "reply_B"]].copy()
    for c in CRITERIA:
        sheet[c] = ""
    sheet["comments"] = ""
    sheet.to_csv(os.path.join(a.outdir, "human_pairwise_sheet_BLINDED.csv"), index=False)
    for i in range(a.n_annotators):
        sheet.to_csv(os.path.join(a.outdir, f"human_pairwise_sheet_BLINDED_annotator{i + 1}.csv"), index=False)

    print(f"wrote {len(sheet)} blinded items "
          f"({df['contrast'].value_counts().to_dict()}) + private key + "
          f"{a.n_annotators} annotator copies -> {a.outdir}")
    print("Give annotators ONLY human_pairwise_sheet_BLINDED*.csv and ANNOTATION_GUIDELINES.md.")


if __name__ == "__main__":
    main()
