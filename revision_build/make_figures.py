"""Regenerate the figures whose values changed, using ONLY the verified CSVs in tables.zip."""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

T = "/projects/sandbox/GCS/_work/tables"
OUT = "/projects/sandbox/GCS/UPDATED_FIGURES"
os.makedirs(OUT, exist_ok=True)

overall = pd.read_csv(f"{T}/overall_A7.csv")
base = pd.read_csv(f"{T}/baseline_comparison.csv")
abl = pd.read_csv(f"{T}/ablation_study.csv")
lens = pd.read_csv(f"{T}/response_lengths.csv")
sota = pd.read_csv(f"{T}/sota_designA.csv")
stats = pd.read_csv(f"{T}/statistical_contrasts_bertscore.csv")

# ---------------- Figure 4: overall performance of A7 (horizontal bars) ----------------
labels = ["BERTScore-F1", "ROUGE-L F1", "Input Relevance", "Grounding Support",
          "Defects per Response", "Latency (s)"]
keys = ["BERTScore-F1", "ROUGE-L F1", "Input Relevance", "Grounding Support (\u03c4=0.50)",
        "Defects per Response", "Latency (s)"]
vals = [float(overall.loc[overall.metric == k, "value"].iloc[0]) for k in keys]
fig, ax = plt.subplots(figsize=(9.2, 4.72))
y = range(len(labels))
ax.barh(list(y), vals, color="#1f77b4")
ax.set_yticks(list(y)); ax.set_yticklabels(labels)
ax.invert_yaxis()
for i, v in enumerate(vals):
    ax.text(v + 0.03, i, f"{v:.3f}".rstrip("0").rstrip(".") if v >= 1 else f"{v:.3f}",
            va="center", fontsize=10)
ax.set_xlim(0, 2.8)
ax.set_xlabel("Reported value")
ax.set_title("Overall Performance of the Proposed PsyAdapt Framework (A7)")
ax.grid(axis="x", linestyle="--", alpha=0.4)
fig.tight_layout()
fig.savefig(f"{OUT}/figure4_overall_performance_A7.png", dpi=200)
plt.close(fig)

# ---------------- Figure 5: quality (defects) vs latency ----------------
allcfg = pd.concat([base[["config", "defects", "latency"]], abl[["config", "defects", "latency"]]]
                   ).drop_duplicates("config").set_index("config")
order = ["A0", "A1", "A2", "A3", "A4", "A5", "A6", "A7"]
fig, ax = plt.subplots(figsize=(7.2, 4.98))
for c in order:
    ax.scatter(allcfg.loc[c, "latency"], allcfg.loc[c, "defects"], s=120, label=c)
ax.set_xlabel("LLM latency per response (s, batch=1)")
ax.set_ylabel("defects per response (lower better)")
ax.set_title("Response Quality vs. Latency (PsyAdapt)")
ax.legend(loc="upper right")
fig.tight_layout()
fig.savefig(f"{OUT}/figure5_quality_vs_latency.png", dpi=200)
plt.close(fig)

# ---------------- Figure 7: response length (median, IQR, mean) ----------------
lens = lens.set_index("config").loc[order]
fig, ax = plt.subplots(figsize=(9.0, 4.05))
for i, c in enumerate(order):
    q1, med, q3, mean = lens.loc[c, ["q1", "median", "q3", "mean"]]
    ax.add_patch(plt.Rectangle((i - 0.25, q1), 0.5, q3 - q1, facecolor="white", edgecolor="black"))
    ax.plot([i - 0.25, i + 0.25], [med, med], color="#ff7f0e", lw=2)
    ax.plot([i], [mean], marker="D", color="#1f77b4", ms=5)
ax.set_xlim(-0.6, len(order) - 0.4)
ax.set_ylim(80, 230)
ax.set_xticks(range(len(order))); ax.set_xticklabels(order)
ax.set_ylabel("Response length (words)")
ax.set_title("Response length by configuration (median, interquartile range, mean)")
ax.plot([], [], color="#ff7f0e", lw=2, label="median")
ax.plot([], [], marker="D", color="#1f77b4", ls="none", ms=5, label="mean")
ax.legend(loc="lower right", fontsize=8)
fig.tight_layout()
fig.savefig(f"{OUT}/figure7_response_lengths.png", dpi=200)
plt.close(fig)

# ---------------- Figure 8: SOTA DESIGN A with bootstrap CIs ----------------
s = sota.copy()
s["short"] = s["arm"].str.replace("S7_", "full ctx: ", regex=False).str.replace("S0_", "bare: ", regex=False)
s = s.sort_values("bertscore")
fig, ax = plt.subplots(figsize=(9.0, 4.4))
err = [(s["bertscore"] - s["bs_lo"]).values, (s["bs_hi"] - s["bertscore"]).values]
colors = ["#1f77b4" if a.startswith("S7") else "#9ecae1" for a in s["arm"]]
ax.bar(s["short"], s["bertscore"], yerr=err, capsize=4, color=colors)
ax.set_ylim(0.55, 0.66)
ax.set_ylabel("BERTScore-F1 (mean, 95% bootstrap CI)")
ax.set_title("Generator swap inside the identical PsyAdapt pipeline (DESIGN A)")
plt.setp(ax.get_xticklabels(), rotation=25, ha="right", fontsize=8)
ax.grid(axis="y", linestyle="--", alpha=0.4)
fig.tight_layout()
fig.savefig(f"{OUT}/figure8_sota_design_a.png", dpi=200)
plt.close(fig)

# ---------------- Figure 9: paired contrasts (forest plot) ----------------
st = stats.copy().iloc[::-1]
fig, ax = plt.subplots(figsize=(9.0, 4.2))
ypos = range(len(st))
for i, (_, r) in zip(ypos, st.iterrows()):
    sig = r["holm_p"] < 0.05
    ax.plot([r["ci_lo"], r["ci_hi"]], [i, i], color="black", lw=1.4)
    ax.plot([r["mean_diff"]], [i], marker="o", ms=8,
            color="#1f77b4" if sig else "#bbbbbb")
ax.axvline(0, color="red", lw=1, ls="--")
ax.set_yticks(list(ypos)); ax.set_yticklabels(st["pair"])
ax.set_xlabel("Paired mean difference in BERTScore-F1 (95% bootstrap CI)")
ax.set_title("Paired contrasts against the full configuration (n = 300 items)")
ax.plot([], [], marker="o", ls="none", color="#1f77b4", label="significant after Holm correction")
ax.plot([], [], marker="o", ls="none", color="#bbbbbb", label="not significant after Holm correction")
ax.legend(loc="lower right", fontsize=8)
ax.grid(axis="x", linestyle="--", alpha=0.4)
fig.tight_layout()
fig.savefig(f"{OUT}/figure9_paired_contrasts.png", dpi=200)
plt.close(fig)

print("figures written:", sorted(os.listdir(OUT)))
