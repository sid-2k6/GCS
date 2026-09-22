"""Final consistency audit: every table cell in the revised paper vs the verified CSVs."""
import pandas as pd
from docx import Document

T = "/projects/sandbox/GCS/_work/tables"
d = Document("/projects/sandbox/GCS/UPDATED_REFERENCE_PAPER.docx")


def get_table(first_cell, must_contain=None):
    for t in d.tables:
        if t.rows[0].cells[0].text.strip() == first_cell:
            if must_contain is None or any(must_contain in c.text for r in t.rows for c in r.cells):
                return t
    raise KeyError(first_cell)


def rows_of(t):
    return {r.cells[0].text.strip(): [c.text.strip() for c in r.cells[1:]] for r in t.rows[1:]}


fails = []


def chk(label, got, exp):
    if str(got).strip() != str(exp).strip():
        fails.append(f"{label}: paper={got!r} expected={exp!r}")


# ---- Table 8 vs baseline_comparison.csv ----
base = pd.read_csv(f"{T}/baseline_comparison.csv").set_index("config")
t8 = rows_of(get_table("Configuration", "A0 \u2013 Generic LLM"))
keymap = {"A0": "A0 \u2013 Generic LLM", "A1": "A1 \u2013 LLM + Standard RAG",
          "A2": "A2 \u2013 RAG + User Profile + Classifier Strategy", "A7": "A7 \u2013 Full PsyAdapt"}
cols = ["bertscore", "rougeL", "relevance", "grounding", "defects", "latency"]
for cfg, rowname in keymap.items():
    got = t8[rowname]
    for j, col in enumerate(cols):
        exp = base.loc[cfg, col]
        if pd.isna(exp):
            chk(f"T8 {cfg} {col}", got[j], "\u2013 (no retrieval)")
        else:
            chk(f"T8 {cfg} {col}", got[j], f"{exp:.3f}" if col not in ("defects", "latency") else f"{exp:.2f}")

# ---- Table 10 vs ablation_study.csv ----
abl = pd.read_csv(f"{T}/ablation_study.csv").set_index("config")
t10 = rows_of(get_table("Configuration", "A3 \u2013 Without Emotion Detection"))
amap = {"A2": "A2 \u2013 RAG + User Profile + Classifier Strategy",
        "A3": "A3 \u2013 Without Emotion Detection", "A4": "A4 \u2013 Without Personality Detection",
        "A5": "A5 \u2013 Without Memory (single-turn control)", "A6": "A6 \u2013 Without Strategy Guidance",
        "A7": "A7 \u2013 Full PsyAdapt"}
for cfg, rowname in amap.items():
    got = t10[rowname]
    for j, col in enumerate(cols):
        exp = abl.loc[cfg, col]
        chk(f"T10 {cfg} {col}", got[j], f"{exp:.3f}" if col not in ("defects", "latency") else f"{exp:.2f}")

# ---- Table 7 vs overall_A7.csv ----
ovr = pd.read_csv(f"{T}/overall_A7.csv").set_index("metric")["value"]
t7 = rows_of(get_table("Evaluation Metric"))
chk("T7 bertscore", t7["BERTScore-F1"][0], ovr["BERTScore-F1"])
chk("T7 rouge", t7["ROUGE-L F1"][0], ovr["ROUGE-L F1"])
chk("T7 relevance", t7["Input Relevance"][0], ovr["Input Relevance"])
chk("T7 grounding", t7["Grounding Support (\u03c4 = 0.50)"][0], ovr["Grounding Support (\u03c4=0.50)"])
chk("T7 meanmax", t7["Mean-max grounding similarity"][0], ovr["Mean-max grounding similarity"])
chk("T7 defects", t7["Defects per Response"][0], ovr["Defects per Response"])
chk("T7 latency", t7["Latency (s)"][0], ovr["Latency (s)"])

# ---- Table 9 vs sota_designA.csv ----
sota = pd.read_csv(f"{T}/sota_designA.csv")
t9 = rows_of(get_table("Arm (generator, context)"))
armmap = {
    "S7_Qwen2.5-3B (A7)": "PsyAdapt A7 \u2013 Qwen2.5-3B-Instruct, full context",
    "S7_Gemma 3 4B": "Gemma 3 4B, full PsyAdapt context",
    "S7_Llama 3.2 3B": "Llama 3.2 3B Instruct, full PsyAdapt context",
    "S7_Mistral 7B": "Mistral 7B Instruct, full PsyAdapt context",
    "S0_Gemma 3 4B (bare)": "Gemma 3 4B, bare",
    "S0_Llama 3.2 3B (bare)": "Llama 3.2 3B Instruct, bare",
    "S0_Mistral 7B (bare)": "Mistral 7B Instruct, bare",
}
for _, r in sota.iterrows():
    rowname = armmap[r["arm"]]
    got = t9[rowname]
    chk(f"T9 {r['arm']} bertscore+CI", got[0], f"{r['bertscore']:.3f} [{r['bs_lo']:.3f}, {r['bs_hi']:.3f}]")
    chk(f"T9 {r['arm']} rouge", got[1], f"{r['rougeL']:.3f}")
    chk(f"T9 {r['arm']} relevance", got[2], f"{r['relevance']:.3f}")
    if pd.isna(r["grounding"]):
        chk(f"T9 {r['arm']} grounding", got[3], "\u2013 (no retrieval)")
    else:
        chk(f"T9 {r['arm']} grounding", got[3], f"{r['grounding']:.3f}")
    chk(f"T9 {r['arm']} defects", got[4], f"{r['defects']:.2f}")
    chk(f"T9 {r['arm']} latency", got[5], f"{r['latency']:.2f}")
# A0 row traced to baseline_comparison.csv
a0 = t9["A0 \u2013 Qwen2.5-3B-Instruct, bare"]
chk("T9 A0 rouge", a0[1], f"{base.loc['A0','rougeL']:.3f}")
chk("T9 A0 defects", a0[4], f"{base.loc['A0','defects']:.2f}")

# ---- Table 11 vs component_classifiers.csv ----
cls = pd.read_csv(f"{T}/component_classifiers.csv")
t11 = rows_of(get_table("Component", "Macro ROC-AUC (OvR)"))
cmap = {"Emotion": "Emotion", "Personality": "Personality", "Strategy v2": "Strategy (deployed, context-only)",
        "Strategy (legacy, leaky)": "Strategy (legacy, label-leaking; diagnostic only)"}
for _, r in cls.iterrows():
    got = t11[cmap[r["component"]]]
    chk(f"T11 {r['component']} acc", got[2], f"{r['accuracy']:.2f}")
    chk(f"T11 {r['component']} macroF1", got[3], f"{r['macro_f1']:.2f}")
    chk(f"T11 {r['component']} wF1", got[4], f"{r['weighted_f1']:.2f}")
    chk(f"T11 {r['component']} auc", got[5], f"{r['roc_auc']:.2f}")

# ---- Table 12 vs statistical_contrasts_bertscore.csv ----
st = pd.read_csv(f"{T}/statistical_contrasts_bertscore.csv")
t12 = rows_of(get_table("Contrast"))
for _, r in st.iterrows():
    name = r["pair"].replace("-", " \u2013 ")
    got = t12[name]
    chk(f"T12 {r['pair']} diff", got[0], f"+{r['mean_diff']:.3f}")
    chk(f"T12 {r['pair']} ci", got[1], f"[{r['ci_lo']:.3f}, {r['ci_hi']:.3f}]")
    chk(f"T12 {r['pair']} dz", got[2], f"{r['d_z']:.2f}")
    sig = "Significant" if r["holm_p"] < 0.05 else "Not significant"
    chk(f"T12 {r['pair']} reading", got[5], sig)

# ---- Table 13 vs memory_experiment.csv ----
mem = pd.read_csv(f"{T}/memory_experiment.csv").iloc[0]
t13 = rows_of(get_table("Metric", "With memory (A2-M)"))
got = t13["BERTScore-F1 (40 ESConv conversations)"]
chk("T13 with", got[0], f"{mem['mean_a2m']:.3f}")
chk("T13 without", got[1], f"{mem['mean_a5m']:.3f}")
chk("T13 diff", got[2], f"+{mem['diff']:.3f}")
chk("T13 ci", got[3], f"[{mem['ci_lo']:.3f}, {mem['ci_hi']:.3f}]")
chk("T13 p", got[5], f"{mem['wilcoxon_p']:.3f}")

# ---- Table 14 vs response_lengths.csv ----
lens = pd.read_csv(f"{T}/response_lengths.csv").set_index("config")
t14 = rows_of(get_table("Configuration", "Median (words)"))
lmap = {"A0": "A0 \u2013 Generic LLM", "A1": "A1 \u2013 LLM + Standard RAG",
        "A2": "A2 \u2013 RAG + User Profile + Classifier Strategy", "A3": "A3 \u2013 Without Emotion Detection",
        "A4": "A4 \u2013 Without Personality Detection", "A5": "A5 \u2013 Without Memory (single-turn control)",
        "A6": "A6 \u2013 Without Strategy Guidance", "A7": "A7 \u2013 Full PsyAdapt"}
for cfg, rowname in lmap.items():
    got = t14[rowname]
    for j, col in enumerate(["median", "q1", "q3", "mean"]):
        chk(f"T14 {cfg} {col}", got[j], str(int(lens.loc[cfg, col])))

# ---- Table 15 / 16 ----
lc = pd.read_csv(f"{T}/length_metric_correlations.csv")
t15 = rows_of(get_table("Metric", "Spearman rho with response length"))
m15 = {"BERTScore-F1": "BERTScore-F1", "ROUGE-L F1": "ROUGE-L F1", "Input relevance": "Input relevance",
       "Grounding support": "Grounding support"}
for _, r in lc.iterrows():
    key = m15[r["metric"].replace("-F1", "-F1").replace("BERTScore-F1", "BERTScore-F1")
              .replace("ROUGE-L F1", "ROUGE-L F1").replace("Input relevance", "Input relevance")
              .replace("Grounding support", "Grounding support")]
    chk(f"T15 {r['metric']} rho", t15[key][0], f"{r['spearman_rho']:.2f}")

dc = pd.read_csv(f"{T}/defect_composition_A7.csv")
t16 = rows_of(get_table("Defect category"))
dmap = {"Unsupported medical claim": "Unsupported medical claim",
        "Overstated certainty": "Overstated certainty",
        "Prescriptive advice w/o disclaimer": "Prescriptive advice without disclaimer",
        "Referral missing (high distress)": "Referral missing (high-distress input)",
        "Diagnostic overreach": "Diagnostic overreach", "Inappropriate boundary": "Inappropriate boundary",
        "TOTAL (A7)": "Total defects per response"}
for _, r in dc.iterrows():
    chk(f"T16 {r['defect_type']}", t16[dmap[r["defect_type"]]][0], f"{r['count_per_response']:.2f}")

print("MISMATCHES:", len(fails))
for f in fails:
    print("  !", f)
