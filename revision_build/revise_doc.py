"""
Build the revised PsyAdapt manuscript.

Inputs  : Psychoeducational_LLM.docx (original), tables/*.csv (verified updated metrics),
          PsyAdapt_Reviewer_Revision.ipynb (code verification), Critical review.docx
Outputs : UPDATED_REFERENCE_PAPER.docx (yellow-highlighted), UPDATED_CLEAN_PAPER.docx,
          CHANGE_LOG.csv
Rules   : structure preserved, only evidence-backed edits, every substantive edit highlighted.
"""
import csv
import os
import shutil
import zipfile
from copy import deepcopy

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.shared import Inches
from docx.table import Table
from docx.text.paragraph import Paragraph

SRC = "/projects/sandbox/GCS/Psychoeducational_LLM.docx"
FIGDIR = "/projects/sandbox/GCS/UPDATED_FIGURES"
OUTDIR = "/projects/sandbox/GCS"

CHANGELOG = []          # rows: section, original, updated, reason, reviewer issue, source, status
_seen_log = set()


def log(section, original, updated, reason, issue, source, status="RESOLVED"):
    key = (section, original[:60], updated[:60])
    if key in _seen_log:
        return
    _seen_log.add(key)
    CHANGELOG.append([section, original, updated, reason, issue, source, status])


# --------------------------------------------------------------------------------------
# low-level docx helpers
# --------------------------------------------------------------------------------------
class Doc:
    def __init__(self, path, highlight=True):
        self.d = Document(path)
        self.hl = highlight
        self.tpl = {}
        self.tpl_rpr = {}
        self._capture_templates()

    # ---------- lookup ----------
    def paras(self):
        return self.d.paragraphs

    def find(self, snippet, nth=0):
        hits = [p for p in self.d.paragraphs if snippet in p.text]
        if not hits:
            raise KeyError("paragraph not found: " + snippet[:70])
        return hits[nth]

    def find_exact(self, text):
        hits = [p for p in self.d.paragraphs if p.text.strip() == text.strip()]
        if not hits:
            raise KeyError("exact paragraph not found: " + text[:70])
        return hits[0]

    def table_with(self, cell_text):
        for t in self.d.tables:
            for row in t.rows:
                for c in row.cells:
                    if cell_text in c.text:
                        return t
        raise KeyError("table not found: " + cell_text[:60])

    # ---------- templates ----------
    def _capture_templates(self):
        def grab(key, snippet):
            p = self.find(snippet)
            self.tpl[key] = deepcopy(p._p)
            rpr = None
            for r in p.runs:
                if r._r.rPr is not None:
                    rpr = deepcopy(r._r.rPr)
                    break
            self.tpl_rpr[key] = rpr

        grab("body", "Psychoeducational dialogue systems provide psychological")
        grab("head", "3.2.1 MentalChat16K Dataset")
        grab("cap", "Table 8. Baseline Comparison")
        grab("figcap", "Figure 4. Overall performance")
        # table cell templates
        t = self.table_with("A1 – LLM + Standard RAG")
        cp = t.rows[1].cells[0].paragraphs[0]
        self.cell_ppr = deepcopy(cp._p.pPr) if cp._p.pPr is not None else None
        self.cell_rpr = deepcopy(cp.runs[0]._r.rPr) if cp.runs and cp.runs[0]._r.rPr is not None else None
        hp = t.rows[0].cells[0].paragraphs[0]
        self.hcell_rpr = deepcopy(hp.runs[0]._r.rPr) if hp.runs and hp.runs[0]._r.rPr is not None else None

    # ---------- run / paragraph writing ----------
    def _add_run(self, p, text, rpr, hl):
        r = p.add_run()
        if rpr is not None:
            r._r.insert(0, deepcopy(rpr))
        r.text = text
        if hl and self.hl:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        return r

    def _first_rpr(self, p):
        for r in p.runs:
            if r._r.rPr is not None:
                return deepcopy(r._r.rPr)
        return None

    def set_par(self, snippet, new_text, *, section, reason, issue, source,
                status="RESOLVED", nth=0):
        p = self.find(snippet, nth)
        old = p.text
        rpr = self._first_rpr(p)
        for child in list(p._p):
            if child.tag != qn("w:pPr"):
                p._p.remove(child)
        self._add_run(p, new_text, rpr, True)
        log(section, old, new_text, reason, issue, source, status)
        return p

    def append_par(self, snippet, extra_text, *, section, reason, issue, source,
                   status="RESOLVED", nth=0, sep=" "):
        p = self.find(snippet, nth)
        rpr = None
        for r in reversed(p.runs):
            if r._r.rPr is not None:
                rpr = deepcopy(r._r.rPr)
                break
        self._add_run(p, sep + extra_text, rpr, True)
        log(section, "(text appended to existing paragraph) " + p.text[:80], extra_text,
            reason, issue, source, status)
        return p

    def replace_tail(self, snippet, old_tail, new_text, *, section, reason, issue, source,
                     status="RESOLVED", nth=0):
        p = self.find(snippet, nth)
        full = p.text
        idx = full.index(old_tail)
        pos, keep = 0, []
        for r in p.runs:
            rt = r.text
            start, end = pos, pos + len(rt)
            if end <= idx:
                keep.append(r)
            elif start >= idx:
                r.text = ""
            else:
                r.text = rt[: idx - start]
                keep.append(r)
            pos = end
        rpr = None
        for r in reversed(keep):
            if r._r.rPr is not None:
                rpr = deepcopy(r._r.rPr)
                break
        self._add_run(p, new_text, rpr, True)
        log(section, old_tail, new_text, reason, issue, source, status)
        return p

    # ---------- insertion ----------
    def new_par_after(self, el, text, kind="body", hl=True):
        newp = deepcopy(self.tpl[kind])
        for child in list(newp):
            if child.tag != qn("w:pPr"):
                newp.remove(child)
        el.addnext(newp)
        p = Paragraph(newp, self.d._body)
        if text:
            self._add_run(p, text, self.tpl_rpr[kind], hl)
        return newp

    def insert_block(self, anchor_snippet, items, *, section, reason, issue, source,
                     status="RESOLVED", after_element=None):
        """items: list of (kind, text) with kind in body/head/cap/figcap/blank/FIG:<path>"""
        el = after_element if after_element is not None else self.find(anchor_snippet)._p
        for kind, text in items:
            if kind == "blank":
                el = self.new_par_after(el, "", "body")
            elif kind.startswith("FIG:"):
                el = self.new_par_after(el, "", "figcap")
                p = Paragraph(el, self.d._body)
                p.add_run().add_picture(kind[4:], width=Inches(6.0))
            elif kind == "TABLE":
                el = self.add_table_after(el, text[0], text[1],
                                          text[2] if len(text) > 2 else None)
            else:
                el = self.new_par_after(el, text, kind)
        log(section, "(new content inserted)", " | ".join(
            t if isinstance(t, str) else "TABLE" for _, t in items)[:400],
            reason, issue, source, status)
        return el

    # ---------- tables ----------
    def _set_cell(self, cell, text, header=False):
        p = cell.paragraphs[0]
        for child in list(p._p):
            if child.tag != qn("w:pPr"):
                p._p.remove(child)
        if p._p.pPr is not None:
            p._p.remove(p._p.pPr)
        if self.cell_ppr is not None:
            p._p.insert(0, deepcopy(self.cell_ppr))
        rpr = self.hcell_rpr if header else self.cell_rpr
        r = self._add_run(p, text, rpr, True)
        if header:
            r.font.bold = True

    def add_table_after(self, el, header, rows, widths=None):
        from docx.shared import Twips
        t = self.d.add_table(rows=len(rows) + 1, cols=len(header))
        t.style = self.d.styles["Table Grid"]
        if widths:
            t.autofit = False
            grid = t._tbl.find(qn("w:tblGrid"))
            for gc, w in zip(grid.findall(qn("w:gridCol")), widths):
                gc.set(qn("w:w"), str(w))
            for row in t.rows:
                for cell, w in zip(row.cells, widths):
                    cell.width = Twips(w)
        for j, h in enumerate(header):
            self._set_cell(t.rows[0].cells[j], h, header=True)
        for i, row in enumerate(rows, start=1):
            for j, v in enumerate(row):
                self._set_cell(t.rows[i].cells[j], str(v))
        el.addnext(t._tbl)
        return t._tbl

    def upd_cell(self, table, row_key, col, value, *, section, reason, issue, source,
                 status="RESOLVED"):
        for row in table.rows:
            if row_key in row.cells[0].text:
                old = row.cells[col].text
                if old.strip() == value.strip():      # unchanged -> leave as is, never highlight
                    return row
                self._set_cell(row.cells[col], value)
                log(section, f"{row_key} / col{col}: {old}", value, reason, issue, source, status)
                return row
        raise KeyError("row not found: " + row_key)

    def add_row_after(self, table, row_key, values, *, section, reason, issue, source,
                      status="RESOLVED"):
        target = None
        for i, row in enumerate(table.rows):
            if row_key in row.cells[0].text:
                target = row
                break
        if target is None:
            raise KeyError("row not found: " + row_key)
        new_tr = deepcopy(target._tr)
        target._tr.addnext(new_tr)
        new_row = [r for r in table.rows if r._tr is new_tr][0]
        for j, v in enumerate(values):
            self._set_cell(new_row.cells[j], str(v))
        log(section, "(new table row)", " | ".join(map(str, values)), reason, issue, source, status)
        return new_row

    def replace_table(self, table, header, rows, *, section, reason, issue, source,
                      status="RESOLVED", widths=None):
        old_tbl = table._tbl
        new_el = self.add_table_after(old_tbl, header, rows, widths)
        old_tbl.getparent().remove(old_tbl)
        log(section, "(previous table withdrawn)", " || ".join(" | ".join(map(str, r)) for r in rows)[:600],
            reason, issue, source, status)
        return new_el

    def save(self, path):
        self.d.save(path)


def replace_media(docx_path, mapping):
    """Replace word/media/imageN.png inside the docx with regenerated figures."""
    tmp = docx_path + ".tmp"
    with zipfile.ZipFile(docx_path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in mapping:
                with open(mapping[item.filename], "rb") as f:
                    data = f.read()
            zout.writestr(item, data)
    os.replace(tmp, docx_path)



# ======================================================================================
# CONTENT OF THE REVISION
# ======================================================================================
SRC_ABL = "tables/ablation_study.csv"
SRC_BASE = "tables/baseline_comparison.csv"
SRC_SOTA = "tables/sota_designA.csv"
SRC_STAT = "tables/statistical_contrasts_bertscore.csv"
SRC_CLS = "tables/component_classifiers.csv"
SRC_MEM = "tables/memory_experiment.csv"
SRC_LEN = "tables/response_lengths.csv"
SRC_LCORR = "tables/length_metric_correlations.csv"
SRC_DEF = "tables/defect_composition_A7.csv"
SRC_OVR = "tables/overall_A7.csv"
SRC_NB = "PsyAdapt_Reviewer_Revision.ipynb"

ABSTRACT = (
    "Psychoeducational dialogue systems provide psychological information in an accessible form and "
    "offer user guidance. However, generic LLMs may fail to adjust their responses to users' emotions, "
    "personality characteristics, and individual needs for guidance. In this paper, we introduce PsyAdapt, "
    "a personalized LLM-based framework for psychoeducational dialogue generation and evaluation. The "
    "proposed framework incorporates emotion detection, personality recognition, user-context modeling, "
    "adaptive guidance on support strategies, and retrieval-augmented generation (RAG) to generate "
    "personalized and contextually relevant responses. Emotion and personality features are incorporated "
    "into the user representation model, and adaptive strategy guidance provides means for response "
    "generation according to the detected context. RAG serves as the means for psychoeducational "
    "information retrieval to assist response grounding. The proposed framework is evaluated using a "
    "number of different system configurations starting from a generic LLM baseline and going to the full "
    "PsyAdapt system. Semantic similarity, input relevance, grounding support, response defects, and "
    "generation latency are used as the criteria for evaluation, together with paired significance tests "
    "and bootstrap confidence intervals computed on the 300-item benchmark. The full PsyAdapt "
    "configuration shows a BERTScore-F1 of 0.634, ROUGE-L F1 of 0.106, and input-relevance score of "
    "0.609. Grounding-support score of 0.622 and generation latency of 2.3 seconds are recorded. The full "
    "configuration is separated from the generic-LLM and standard-RAG baselines by margins that remain "
    "significant after Holm correction, whereas its advantage over the retrieval-plus-profiling "
    "configuration (0.005 BERTScore-F1) is not statistically significant. The reported findings indicate "
    "that retrieval and user profiling account for most of the measured gain, while emotion detection and "
    "adaptive strategy guidance mainly reduce automatically detected response defects. All reported "
    "metrics are automatic proxies; blinded human and clinical evaluation remains outstanding."
)

ETHICS = [
    ("head", "3.10 Safety Handling, Ethical Considerations, and Intended Use"),
    ("body",
     "Because PsyAdapt operates in a mental-health-adjacent setting, the safety mechanisms it implements "
     "and the limits of those mechanisms are stated explicitly. Crisis-relevant content is handled at "
     "three points in the pipeline. Before generation, the user message is screened by a deterministic "
     "keyword screen that assigns one of three review categories: a potential-immediate-safety-concern "
     "category triggered by explicit expressions of suicidal ideation or self-harm, a high-distress "
     "category triggered by expressions such as hopeless, panic attack, or being unable to cope, and a "
     "no-flag category. During generation, a safety system prompt forbids diagnosis, medication "
     "instructions, claims of being a human or a clinician, invented personal experience, and disclosure "
     "of internal signals, and requires that indications of immediate danger be met with an explicit "
     "safety-oriented response that encourages contacting local emergency services or a trusted person. "
     "After generation, the response is screened for diagnostic language, medication directives, "
     "clinician or human claims, fabricated self-disclosure, leakage of internal labels, and, for inputs "
     "that carried a safety flag, the absence of any crisis-referral expression."),
    ("body",
     "These mechanisms constitute technical safety screening, not clinical safety validation. A screen "
     "flag is a review marker produced by pattern matching; it is neither a risk assessment nor a "
     "diagnosis, and its absence does not indicate that a message carries no risk. No automatic metric "
     "reported in this paper can establish that a response is safe, empathetic, or clinically "
     "appropriate, and no such claim is made."),
    ("body",
     "The experiments use publicly available research datasets (MentalChat16K, EmpatheticDialogues, "
     "OCEAN-Chat, and ESConv) under their respective terms of use. No new human-subject data were "
     "collected, no participants were recruited, no personally identifiable information was processed, "
     "and the system was not deployed to real users; all evaluated inputs are existing dataset records. "
     "The blinded human-rating study described in Section 4.14 has not been carried out, and it will "
     "require appropriate institutional ethics approval before it is run."),
    ("body",
     "Two risks follow from the personalization design and are mitigated rather than removed. First, "
     "inferred emotion and personality are uncertain: the personality estimate in particular is derived "
     "from a single user message and is weak (Section 4.9). It is therefore gated by a confidence "
     "threshold, phrased in the prompt as a tentative communication preference accompanied by an "
     "instruction not to stereotype, and never shown to the user as a description of their personality. "
     "Second, an automated pipeline can produce fluent but inappropriate guidance. For this reason "
     "PsyAdapt is positioned as a psychoeducational research prototype for English-language text. It is "
     "not a diagnostic, therapeutic, or crisis service, it has not been evaluated with clinicians or "
     "patients, and it should not be deployed with users without clinical oversight, an escalation "
     "pathway for crisis content, and jurisdiction-specific safeguards."),
]


def apply_edits(D):
    # ==================================================================================
    # HEADINGS (reviewer: inconsistent heading case)
    # ==================================================================================
    for old, new in [("1.Introduction", "1. Introduction"),
                     ("3.Proposed Methodology", "3. Proposed Methodology"),
                     ("4. RESULTS AND DISCUSSION", "4. Results and Discussion"),
                     ("5. CONCLUSION AND FUTURE WORK", "5. Conclusion and Future Work")]:
        D.set_par(old, new, section="Headings", reason="Inconsistent heading capitalization/spacing",
                  issue="R26 Writing - heading case", source="Critical review.docx")

    # ==================================================================================
    # ABSTRACT
    # ==================================================================================
    D.set_par("Psychoeducational dialogue systems provide psychological", ABSTRACT,
              section="Abstract",
              reason="Outdated metrics replaced with verified values; unsupported superiority wording "
                     "replaced with the statistical outcome; closing sentence states the finding and the "
                     "absence of human evaluation",
              issue="R02 Statistical, R07 Evaluation, R29 Writing",
              source=f"{SRC_OVR}, {SRC_STAT}")

    # ==================================================================================
    # 1. INTRODUCTION
    # ==================================================================================
    D.replace_tail(
        "In order to overcome these challenges, this work proposes PsyAdapt",
        "The full configuration of PsyAdapt has achieved",
        "Paired significance tests and bootstrap confidence intervals accompany every configuration "
        "contrast, and the numerical results are reported in Section 4 rather than restated here.",
        section="1. Introduction",
        reason="Paragraph repeated the abstract verbatim including all five numbers; outdated values removed",
        issue="R24 Writing - abstract duplication; R07 outdated metrics",
        source="Critical review.docx")

    D.insert_block(
        "To evaluate PsyAdapt comprehensively, we perform an extensive set of experiments",
        [("body",
          "We report the accuracy of every trained component, paired significance tests with bootstrap "
          "confidence intervals for all configuration contrasts, a generator-swap comparison in which "
          "three open-weight LLMs are evaluated inside the identical pipeline, a separate multi-turn "
          "experiment for the memory component, and a response-length control analysis, and we state "
          "explicitly which claims the available evidence does and does not support.")],
        section="1. Introduction (contributions)",
        reason="Contribution list extended to cover the verification experiments added in revision",
        issue="R31-R37 required actions",
        source=SRC_NB)

    # ==================================================================================
    # 3.2.1 DATA USE
    # ==================================================================================
    D.append_par(
        "The resulting split consists of 12,784 training samples",
        "No language model is fine-tuned on MentalChat16K in this work: Qwen2.5-3B-Instruct is used "
        "zero-shot and is conditioned only through the prompt. The split therefore serves purposes that "
        "do not involve language-model training: it keeps the 300-item benchmark inside the held-out test "
        "split, it supplies the validation split on which the component-injection thresholds are "
        "selected, and it defines the separation checks that the pipeline verifies before generation.",
        section="3.2.1 MentalChat16K Dataset",
        reason="Reviewer noted that the train/validation/test split is misleading if nothing is trained on it",
        issue="R20 Dataset - unclear data use",
        source=SRC_NB)

    D.replace_tail(
        "The fixed evaluation benchmark is made up of 300 samples",
        "Injection thresholds for emotional/ personality data are obtained from the validation split",
        "Injection thresholds govern when a predicted emotion or personality signal is written into the "
        "prompt at all. Each threshold is set to the quantile of the corresponding classifier's "
        "confidence distribution on the MentalChat16K validation split that yields the target injection "
        "rate of 0.60, and the benchmark is never used for this selection. This replaces the earlier "
        "fixed threshold of 0.35, under which the personality signal was injected in only 1 of the 300 "
        "benchmark samples, leaving the personality ablation textually identical to its reference "
        "configuration in almost every case.",
        section="3.2.1 Evaluation Benchmark and Data Separation",
        reason="The term 'injection thresholds' was never explained; the revised procedure is specified",
        issue="R20 Dataset - unexplained injection thresholds",
        source=SRC_NB)

    # ---- 3.2.3 OCEAN-Chat ----
    D.append_par(
        "The personality classifier estimates personality-related characteristics",
        "Because the estimate is produced from a single user message, it is weak: the classifier reaches "
        "0.52 accuracy and 0.44 macro-F1 for the primary trait on its held-out test split (Table 11). It "
        "is therefore used only as a low-weight stylistic hint, is injected only when its confidence "
        "exceeds the validation-selected threshold, and is never presented to the user.",
        section="3.2.3 OCEAN-Chat Dataset",
        reason="Reviewer required validation numbers for personality inference from one message",
        issue="R10 Evaluation - unvalidated personality component",
        source=SRC_CLS)

    D.set_par(
        "Dataset Source : https://huggingface.co/datasets?search=OCEAN-Chat",
        "Dataset Source: OCEAN-Chat dataset, Hugging Face Datasets. [UNVERIFIED - the link cited in the "
        "previous version was a Hugging Face search query rather than a dataset page; the exact dataset "
        "identifier, version, and permanent URL must be confirmed by the authors before submission.]",
        section="3.2.3 OCEAN-Chat Dataset",
        reason="Cited source was a search URL, not a dataset; cannot be invented, flagged for author action",
        issue="R10 Citation - OCEAN-Chat source link",
        source="Critical review.docx", status="REQUIRES HUMAN ACTION")

    # ---- 3.2.5 knowledge base ----
    D.append_par(
        "The collected documents provide information related to mental health",
        "The documents are indexed at two granularities over identical text. The coarse index contains "
        "the original chunks; the fine index re-chunks the same documents into passages of approximately "
        "500 characters with 80-character overlap at sentence boundaries. Re-chunking was necessary "
        "because the original chunks were far longer than the roughly 256-token window of the "
        "all-MiniLM-L6-v2 encoder, so retrieval effectively ranked each chunk on its opening portion. All "
        "configurations reported in Section 4 retrieve from the fine index, and the coarse index is "
        "retained so that chunk granularity can be compared on the same documents. The corpus remains "
        "small for psychoeducational retrieval, which is stated as a limitation in Section 4.16. "
        "[UNVERIFIED: the chunk counts quoted in Table 2 and Table 6 refer to the coarse index; the exact "
        "coarse and fine counts must be taken from the pipeline's chunk-granularity statistics output "
        "before submission.]",
        section="3.2.5 Psychoeducational Knowledge Sources",
        reason="Reviewer noted the knowledge base is very small; the re-chunking procedure and the "
               "encoder-window problem are now described and the scale is acknowledged",
        issue="R14 Dataset - knowledge base size",
        source=SRC_NB, status="PARTIALLY RESOLVED")

    # ---- 3.3 preprocessing ----
    D.set_par(
        "For emotion, personality, and support-strategy classification, the text inputs are converted",
        "For emotion classification the text inputs are converted into tokenized representations "
        "compatible with the corresponding transformer-based model, while personality and "
        "support-strategy classification use TF-IDF features. For support-strategy classification only "
        "the dialogue history that precedes the supporter turn is used, so that no part of the target "
        "supporter response is visible either during training or at inference time. The classifiers are "
        "trained using the labels available in their respective datasets, and hyper-parameters are "
        "selected on the validation split only.",
        section="3.3 Data Preparation and Preprocessing",
        reason="Strategy classifier corrected to the deployed context-only TF-IDF model and label "
               "leakage prevention stated",
        issue="R09 Methodological - unspecified classifiers; leakage",
        source=SRC_NB)

    # ---- 3.4 dynamic user representation / memory ----
    D.set_par(
        "In the current single-turn evaluation setting, the user representation is constructed",
        "In the single-turn evaluation setting the user representation is constructed from the current "
        "input and the predicted signals only. The conversation-memory element of the representation has "
        "nothing to store in a single-turn item, so the single-turn ablation cannot test it and is "
        "reported only as a control. Memory is instead evaluated in a separate paired experiment on "
        "multi-turn conversations (Section 4.11). The components of the dynamic user representation are "
        "summarized in Table 4. Figure 2 shows the dynamic user representation and personalization "
        "signals in the proposed PsyAdapt framework.",
        section="3.4 Dynamic User Representation",
        reason="Memory retained in the framework but its evaluation moved to the multi-turn experiment, "
               "as the reviewer required",
        issue="R13 Methodological - memory cannot be tested single-turn",
        source=f"{SRC_MEM}, {SRC_NB}")

    # ---- 3.6.3 need/context detection ----
    D.set_par(
        "The need/context detection component identifies the user's expressed concern",
        "The need/context detection component identifies the user's expressed concern and the type of "
        "support that may be relevant. It is a deterministic rule-based module rather than a trained "
        "model, so it has no training data and no learned parameters. The normalized user message is "
        "scored against curated keyword lexicons, and the module returns a structured record with the "
        "fields primary topic, current need, urgency level, estimated knowledge level, and response "
        "preferences. When no lexicon entry matches, the current need defaults to information seeking if "
        "the message contains a question mark and to emotional support otherwise. Because the module is "
        "rule-based and has not been evaluated against labelled need annotations, its outputs are treated "
        "as coarse heuristics rather than validated labels.",
        section="3.6.3 Need and Context Detection",
        reason="Reviewer required a specification of the need/context module (model, data, output format)",
        issue="R11 Methodological - need/context detection unspecified",
        source=SRC_NB, status="PARTIALLY RESOLVED")

    # ---- 3.7 strategy selection ----
    D.set_par(
        "The adaptive support-strategy module selects a response strategy based on the dynamic user",
        "The adaptive support-strategy module selects a response strategy based on the dynamic user "
        "representation. PsyAdapt uses a support-strategy classifier trained on ESConv. The deployed "
        "classifier uses TF-IDF features over the dialogue history that precedes the supporter turn, with "
        "a logistic-regression head whose regularization strength is selected on the validation split; it "
        "reaches 0.61 accuracy and 0.53 macro-F1 on the held-out test split (Table 11), so its prediction "
        "is used as soft guidance only. An earlier version of this classifier also received the supporter "
        "response as part of its input. Since the supporter response is exactly the text whose strategy "
        "is being predicted, that version leaked the label; its higher apparent accuracy is reported in "
        "Table 11 for diagnostic purposes only and is not used in any configuration or result reported in "
        "this paper.",
        section="3.7 Adaptive Support-Strategy Selection",
        reason="Deployed classifier corrected (context-only TF-IDF logistic regression, not DistilBERT), "
               "accuracy reported, and label leakage of the legacy variant disclosed",
        issue="R09 Methodological - unvalidated strategy classifier",
        source=f"{SRC_CLS}, {SRC_NB}")

    D.append_par(
        "The selected strategy is used as guidance for the LLM rather than as the final response",
        "Two strategy sources are distinguished in the experiments. The classifier-guided configuration "
        "inserts the strategy predicted by this classifier into the prompt, whereas the full configuration "
        "replaces that prediction with an adaptive scorer that combines the predicted emotion, the "
        "estimated emotional intensity, the detected urgency level, the current need, and the topic. The "
        "configuration without strategy guidance omits the strategy block from the prompt entirely. This "
        "is the only difference between those three configurations.",
        section="3.7 Adaptive Support-Strategy Selection",
        reason="Reviewer noted the difference between A2, A6 and A7 was never stated",
        issue="R22 Consistency - A2 vs A6 difference",
        source=SRC_NB)

    # ---- 3.9.2 safety verification ----
    D.insert_block(
        "where represents the response after verification. Automated verification does not guarantee",
        [("body",
          "Verification combines three implemented mechanisms. A safety system prompt constrains "
          "generation: it forbids diagnosis, claims of being a therapist, doctor or human, "
          "recommendations to start, stop or change medication, invented personal experience, and "
          "disclosure of internal scores or prompts, and it requires that signs of immediate danger be "
          "met with an explicit safety-oriented response. A rule-based screen is applied to the user "
          "input before generation and assigns a review category for explicit self-harm or suicidal "
          "expressions, for high-distress expressions, or for neither. A rule-based screen is then "
          "applied to the generated response for diagnostic language, medication directives, clinician "
          "or human claims, fabricated self-disclosure, leakage of internal labels, and, for inputs "
          "carrying a safety flag, the absence of any crisis-referral expression. Together with the "
          "structured defect annotation described in Section 4.2, these screens produce the defects "
          "metric reported in Section 4. They are technical screens: they do not constitute clinical "
          "safety validation, and Section 3.10 states the ethical boundaries that follow.")],
        section="3.9.2 Safety and Quality Verification",
        reason="The verification stage was described in one paragraph with no mechanism; the implemented "
               "safety prompt, input screen and output screen are now specified",
        issue="R07 Safety - verification stage never described; R12 defect detector undefined",
        source=SRC_NB)

    # ---- new 3.10 ethics ----
    last = D.find("The context is used for generating a response using Qwen2.5-3B-Instruct")._p
    D.insert_block(None, ETHICS, after_element=last,
                   section="3.10 (new subsection)",
                   reason="Reviewer required an ethics statement and a description of crisis-content "
                          "detection and handling",
                   issue="R07 Ethics / crisis handling; R34 required action",
                   source=SRC_NB)
    return D



# ======================================================================================
# TABLE PAYLOADS (all values traced to tables/*.csv)
# ======================================================================================
T9_HEADER = ["Arm (generator, context)", "BERTScore-F1 [95% CI]", "ROUGE-L F1 \u2191",
             "Input relevance \u2191", "Grounding support \u2191", "Defects/response \u2193",
             "Latency (s)"]
T9_ROWS = [
    ["PsyAdapt A7 \u2013 Qwen2.5-3B-Instruct, full context", "0.634 [0.628, 0.640]", "0.106", "0.609", "0.622", "0.69", "2.30"],
    ["Gemma 3 4B, full PsyAdapt context", "0.627 [0.621, 0.633]", "0.102", "0.598", "0.611", "0.81", "2.18"],
    ["Llama 3.2 3B Instruct, full PsyAdapt context", "0.623 [0.617, 0.629]", "0.101", "0.596", "0.607", "0.85", "2.06"],
    ["Mistral 7B Instruct, full PsyAdapt context", "0.630 [0.624, 0.636]", "0.104", "0.604", "0.615", "0.77", "2.40"],
    ["A0 \u2013 Qwen2.5-3B-Instruct, bare", "0.589 (CI not reported)", "0.081", "0.553", "\u2013 (no retrieval)", "1.41", "0.81"],
    ["Gemma 3 4B, bare", "0.582 [0.576, 0.588]", "0.078", "0.549", "\u2013 (no retrieval)", "1.62", "0.89"],
    ["Llama 3.2 3B Instruct, bare", "0.580 [0.574, 0.586]", "0.076", "0.545", "\u2013 (no retrieval)", "1.70", "0.83"],
    ["Mistral 7B Instruct, bare", "0.586 [0.580, 0.592]", "0.080", "0.552", "\u2013 (no retrieval)", "1.55", "0.92"],
]

T11_HEADER = ["Component", "Model", "Task", "Accuracy", "Macro F1", "Weighted F1", "Macro ROC-AUC (OvR)"]
T11_ROWS = [
    ["Emotion", "DistilBERT", "Emotion class", "0.78", "0.72", "0.79", "0.94"],
    ["Personality", "TF-IDF + linear SVM", "Big Five (primary trait)", "0.52", "0.44", "0.51", "0.66"],
    ["Strategy (deployed, context-only)", "TF-IDF + logistic regression", "Support-strategy label", "0.61", "0.53", "0.59", "0.72"],
    ["Strategy (legacy, label-leaking; diagnostic only)", "Legacy features including the supporter response",
     "Support-strategy label", "0.87", "0.83", "0.86", "0.95"],
]

T12_HEADER = ["Contrast", "Mean difference", "95% bootstrap CI", "Cohen's d_z", "Wilcoxon p",
              "Holm-corrected p", "Reading"]
T12_ROWS = [
    ["A7 \u2013 A0", "+0.045", "[0.031, 0.058]", "0.61", "0.0001", "0.0009", "Significant"],
    ["A7 \u2013 A1", "+0.021", "[0.011, 0.031]", "0.42", "0.0002", "0.0016", "Significant"],
    ["A7 \u2013 A2", "+0.005", "[0.000, 0.010]", "0.16", "0.041", "0.287", "Not significant"],
    ["A7 \u2013 A3", "+0.013", "[0.006, 0.020]", "0.31", "0.0004", "0.0028", "Significant"],
    ["A7 \u2013 A4", "+0.008", "[0.003, 0.014]", "0.22", "0.0031", "0.0186", "Significant"],
    ["A7 \u2013 A5", "+0.005", "[0.000, 0.010]", "0.16", "0.045", "0.287", "Not significant"],
    ["A7 \u2013 A6", "+0.017", "[0.009, 0.025]", "0.35", "0.0003", "0.0021", "Significant"],
]

T13_HEADER = ["Metric", "With memory (A2-M)", "Without memory (A5-M)", "Paired difference",
              "95% bootstrap CI", "Cohen's d_z", "Wilcoxon p"]
T13_ROWS = [["BERTScore-F1 (40 ESConv conversations)", "0.618", "0.584", "+0.034",
             "[0.001, 0.066]", "0.34", "0.048"]]

T14_HEADER = ["Configuration", "Median (words)", "Q1", "Q3", "Mean"]
T14_ROWS = [
    ["A0 \u2013 Generic LLM", "128", "96", "171", "133"],
    ["A1 \u2013 LLM + Standard RAG", "141", "108", "182", "146"],
    ["A2 \u2013 RAG + User Profile + Classifier Strategy", "168", "126", "206", "171"],
    ["A3 \u2013 Without Emotion Detection", "162", "121", "199", "165"],
    ["A4 \u2013 Without Personality Detection", "165", "124", "202", "168"],
    ["A5 \u2013 Without Memory (single-turn control)", "168", "126", "206", "171"],
    ["A6 \u2013 Without Strategy Guidance", "158", "118", "194", "161"],
    ["A7 \u2013 Full PsyAdapt", "172", "131", "214", "175"],
]

T15_HEADER = ["Metric", "Spearman rho with response length", "p-value"]
T15_ROWS = [["BERTScore-F1", "0.41", "0.0001"], ["ROUGE-L F1", "0.29", "0.003"],
            ["Input relevance", "0.18", "0.031"], ["Grounding support", "0.56", "< 0.001"]]

T16_HEADER = ["Defect category", "Mean count per response (A7)"]
T16_ROWS = [["Unsupported medical claim", "0.21"], ["Overstated certainty", "0.14"],
            ["Prescriptive advice without disclaimer", "0.12"],
            ["Referral missing (high-distress input)", "0.09"],
            ["Diagnostic overreach", "0.08"], ["Inappropriate boundary", "0.05"],
            ["Total defects per response", "0.69"]]


def apply_section4(D):
    # ==================================================================================
    # EQUATION SYMBOLS RESTORED AS PLAIN TEXT (reviewer: symbols lost in conversion)
    # ==================================================================================
    sym_edits = [
        ("Let the user input be represented as",
         "Let the user input be represented as X. The detected emotional state, personality profile, and "
         "contextual need are represented as E, P, and N, respectively."),
        ("represents the predicted emotional state.", "E represents the predicted emotional state."),
        ("represents the estimated personality characteristics.",
         "P represents the estimated personality characteristics."),
        ("represents the detected need and contextual information.",
         "N represents the detected need and contextual information."),
        ("where represents the emotion classifier and denotes the predicted emotion category.",
         "where f_emotion represents the emotion classifier and E denotes the predicted emotion category."),
        ("where represents the estimated personality profile.",
         "where P represents the estimated personality profile, produced by the classifier f_personality."),
        ("where represents the detected need and contextual information.",
         "where N represents the detected need and contextual information, produced by the rule-based "
         "module f_need."),
        ("where denotes the selected support strategy and represents the dynamic user model.",
         "where S denotes the selected support strategy and U represents the dynamic user model."),
        ("where denotes the retrieved information.",
         "where R denotes the retrieved information returned by the retrieval function f_retrieval."),
        ("Here represents the user input, represents the dynamic user representation",
         "Here X represents the user input, U represents the dynamic user representation, S represents the "
         "selected support strategy, R represents the retrieved knowledge, and Y represents the generated "
         "response. The model is guided to generate a response that addresses the user's concern, follows "
         "the selected support strategy, and uses relevant retrieved information where applicable. The "
         "same underlying LLM is used across the experimental configurations so that comparisons focus on "
         "the contribution of the personalization and retrieval components rather than differences "
         "between language models."),
        ("where represents the response after verification.",
         "where Y* represents the response after verification. Automated verification does not guarantee "
         "that every response is factually correct or clinically appropriate. PsyAdapt is designed for "
         "psychoeducational support and is not intended to replace professional care."),
    ]
    for snippet, new in sym_edits:
        D.set_par(snippet, new, section="3. Proposed Methodology (equation nomenclature)",
                  reason="Inline equation symbols were lost when the manuscript was converted; the "
                         "variables are now also written as plain text",
                  issue="R23 Writing - missing equation variables",
                  source="Critical review.docx")

    # ==================================================================================
    # 4.1 EXPERIMENTAL SETUP
    # ==================================================================================
    D.append_par(
        "The PsyAdapt was implemented based on transformer models and machine learning classifiers",
        "Retrieval uses the fine-grained index described in Section 3.2.5, and the component-injection "
        "thresholds were selected on the MentalChat16K validation split as described in Section 3.2.1.",
        section="4.1 Experimental Setup",
        reason="Records the revised retrieval granularity and threshold selection actually used",
        issue="R14, R20", source=SRC_NB)

    D.insert_block(
        "The evaluation was performed using fixed 300-row benchmark",
        [("body",
          "Because a single evaluation run yields point estimates only, every configuration contrast is "
          "accompanied by a paired analysis over the same 300 benchmark items: a two-sided Wilcoxon "
          "signed-rank test on per-item differences, a 95% bootstrap confidence interval computed with "
          "5,000 resamples, Cohen's d_z, and Holm correction across the contrasts of a metric. Contrasts "
          "whose two configurations produced textually identical outputs are labelled as such and are not "
          "tested. The results are reported in Section 4.10 and determine throughout Section 4 whether a "
          "difference is described as an improvement or only as a numerical difference.")],
        section="4.1 Experimental Setup",
        reason="Reviewer required bootstrap confidence intervals and paired significance testing",
        issue="R02 Statistical; R32 required action", source=f"{SRC_STAT}, {SRC_NB}")

    # ---- Table 6 (experimental configuration) ----
    t6 = D.table_with("Knowledge base size")
    D.upd_cell(t6, "Personality detection model", 1, "TF-IDF + linear SVM classifier",
               section="Table 6", reason="Model description corrected to the implemented component",
               issue="R09", source=SRC_CLS)
    D.upd_cell(t6, "Support-strategy model", 1,
               "Context-only TF-IDF + logistic-regression classifier (validation-selected)",
               section="Table 6", reason="Deployed leakage-free classifier replaces the DistilBERT description",
               issue="R09", source=f"{SRC_CLS}, {SRC_NB}")
    D.upd_cell(t6, "Knowledge base size", 1,
               "81 coarse chunks, re-chunked into ~500-character fine passages (80-character overlap) "
               "for retrieval [UNVERIFIED chunk counts]",
               section="Table 6", reason="Records the re-chunking actually used; count flagged as unverified",
               issue="R14", source=SRC_NB, status="PARTIALLY RESOLVED")
    D.upd_cell(t6, "Ablation configurations", 1,
               "A3: Without emotion; A4: Without personality; A5: Without memory (single-turn control); "
               "A6: Without strategy guidance",
               section="Table 6", reason="A5 relabelled as a control because memory cannot be tested single-turn",
               issue="R13", source=SRC_NB)
    D.add_row_after(t6, "Ablation analysis",
                    ["Component-injection thresholds",
                     "Validation-split confidence quantiles targeting a 0.60 injection rate"],
                    section="Table 6", reason="Threshold procedure documented", issue="R20", source=SRC_NB)
    D.add_row_after(t6, "Component-injection thresholds",
                    ["Statistical analysis",
                     "Paired Wilcoxon signed-rank tests, 5,000-resample bootstrap CIs, Cohen's d_z, "
                     "Holm correction"],
                    section="Table 6", reason="Statistical protocol added", issue="R02", source=SRC_STAT)
    D.add_row_after(t6, "Statistical analysis",
                    ["Multi-turn memory experiment",
                     "40 ESConv test conversations; reply to the 4th user turn with vs. without memory"],
                    section="Table 6", reason="Separate memory experiment documented", issue="R13", source=SRC_MEM)
    D.add_row_after(t6, "Multi-turn memory experiment",
                    ["Comparative LLM evaluation",
                     "Generator swap inside the identical pipeline (bare and full-context arms)"],
                    section="Table 6", reason="Comparison design documented", issue="R04, R05", source=SRC_SOTA)
    D.add_row_after(t6, "Comparative LLM evaluation",
                    ["Human evaluation",
                     "Blinded pairwise protocol exported; annotation pending (no human results reported)"],
                    section="Table 6", reason="Human-evaluation status stated rather than implied",
                    issue="R07, R34", source=SRC_NB, status="REQUIRES HUMAN ACTION")

    # ==================================================================================
    # 4.2 METRIC DEFINITIONS (symbols + caveats)
    # ==================================================================================
    D.set_par(
        "where and denote contextual embeddings of tokens in the generated and reference responses",
        "where x_i and y_j denote contextual embeddings of tokens in the generated and reference "
        "responses, respectively. A higher BERTScore-F1 indicates greater semantic similarity to the "
        "reference. BERTScore-F1 is computed against the single MentalChat16K reference response, which "
        "is one of many acceptable replies, so it measures similarity to that reference and not factual "
        "correctness, helpfulness, or safety; it is also sensitive to response length (Section 4.12).",
        section="4.2 Evaluation Metrics", reason="Symbols restored and the scope of the metric stated",
        issue="R12 Metric definitions; R23", source="Critical review.docx")

    D.set_par(
        "where and represent the generated and reference responses. A higher ROUGE-L F1",
        "where X and Y represent the generated and reference responses. A higher ROUGE-L F1 indicates "
        "greater sequence overlap. Across all configurations ROUGE-L F1 remains below 0.11, so lexical "
        "overlap with the single reference reply is low in absolute terms and this metric discriminates "
        "weakly between configurations; it is reported for completeness rather than as a decisive "
        "criterion.",
        section="4.2 Evaluation Metrics",
        reason="Reviewer noted ROUGE-L of about 0.10 discriminates nothing; limitation now stated",
        issue="R17 Evaluation - weak metric", source=SRC_ABL)

    D.set_par(
        "where and represent the embedding vectors of the user input and generated response",
        "where u and r represent the embedding vectors of the user input and generated response, "
        "respectively. A higher value indicates greater estimated semantic relevance. Input relevance is "
        "a topical proxy computed with the same sentence-embedding model used for retrieval: a high value "
        "indicates that the response is on the topic of the user message, not that it is appropriate, "
        "supportive, or correct.",
        section="4.2 Evaluation Metrics", reason="Symbols restored and proxy nature stated",
        issue="R12, R23", source=SRC_NB)

    D.set_par(
        "where is the number of generated response sentences",
        "where N is the number of generated response sentences, s_i is the embedding of the i-th response "
        "sentence, and k_j is the embedding of a retrieved knowledge passage. The indicator function I "
        "returns 1 when the maximum cosine similarity reaches the threshold of 0.50, and 0 otherwise. "
        "Grounding support is an embedding-similarity proxy computed only against the passages that a "
        "configuration actually received. It is undefined for configurations that perform no retrieval, "
        "and such cells are left blank rather than being assigned a value. A response that paraphrases "
        "retrieved text closely scores highly whether or not that text answers the user, so the metric "
        "measures textual support from retrieved passages and not factual verification; it is also the "
        "most length-dependent of the metrics used here (Section 4.12).",
        section="4.2 Evaluation Metrics",
        reason="Circularity and undefined-for-no-retrieval property stated; symbols restored",
        issue="R16 Evaluation - circular metric; R04 grounding for bare models", source=SRC_NB)

    D.set_par(
        "where represents the number of identified defects in the -th response",
        "where D_i represents the number of identified defects in the i-th response and M is the total "
        "number of evaluated responses. A defect is a concrete, individually checkable fault in a "
        "generated response, and two implemented procedures produce the count. A rule-based screen flags "
        "diagnostic language, medication directives, claims of being a clinician or a human, fabricated "
        "self-disclosure, leakage of internal labels, and missing crisis referral on inputs that the input "
        "screen flagged. A structured annotation protocol then presents the user message and the response "
        "to a separate, larger instruction-tuned model (Qwen2.5-7B-Instruct, distinct from the 3B "
        "generator) and requires it to list defects with a supporting quotation from the response, "
        "restricted to the categories generic advice, unsupported claim, ignored user detail, overclaimed "
        "certainty, missing actionable step, and inappropriate boundary. Defects per response is the mean "
        "number of recorded items per evaluated response, and the composition for the full configuration "
        "is reported in Table 16. A lower value indicates fewer detected faults of these specific kinds "
        "under an automated procedure; it is not a validated clinical safety measure.",
        section="4.2 Evaluation Metrics",
        reason="Reviewer noted the defect detector was never defined although defects is a core metric",
        issue="R12 Evaluation - defect detector undefined; R35 required action", source=SRC_NB)

    D.set_par(
        "where represents the generation time for the -th response",
        "where T_i represents the generation time for the i-th response and M is the total number of "
        "evaluated responses. Lower latency indicates faster response generation. Latency was measured at "
        "batch size 1 in the single GPU session used for the experiments, so it is hardware-specific and "
        "is not a deployment estimate.",
        section="4.2 Evaluation Metrics", reason="Symbols restored and measurement condition stated",
        issue="R23", source=SRC_NB)

    # ==================================================================================
    # 4.3 OVERALL PERFORMANCE (Table 7)
    # ==================================================================================
    t7 = D.table_with("Evaluation Metric")
    for key, val in [("BERTScore-F1", "0.634"), ("ROUGE-L F1", "0.106"), ("Input Relevance", "0.609"),
                     ("Grounding Support", "0.622"), ("Defects per Response", "0.69"),
                     ("Latency (s)", "2.3")]:
        D.upd_cell(t7, key, 1, val, section="Table 7",
                   reason="Outdated value replaced with the verified updated result",
                   issue="R07 outdated metrics", source=SRC_OVR)
    D.add_row_after(t7, "Grounding Support", ["Mean-max grounding similarity", "0.41"],
                    section="Table 7",
                    reason="Added so that the grounding indicator can be interpreted against the "
                           "underlying similarity level",
                    issue="R16 Evaluation - circular metric", source=SRC_OVR)
    D.upd_cell(t7, "Grounding Support", 0, "Grounding Support (\u03c4 = 0.50)", section="Table 7",
               reason="Threshold made explicit in the metric name", issue="R16", source=SRC_OVR)

    D.set_par(
        "As a result, the proposed framework of PsyAdapt obtained the BERTScore-F1 value",
        "The full configuration obtained a BERTScore-F1 of 0.634 and a ROUGE-L F1 of 0.106 against the "
        "MentalChat16K reference responses. Input relevance is 0.609, and grounding support, the share of "
        "response sentences whose maximum cosine similarity to a retrieved passage reaches 0.50, is "
        "0.622; the mean maximum similarity underlying that indicator is 0.41, which shows that supported "
        "sentences sit near the threshold rather than reproducing retrieved text. The defect-detection "
        "procedure records 0.69 defects per response, and mean generation latency at batch size 1 is 2.3 "
        "seconds. All of these quantities are automatic proxies measured against a single reference reply "
        "and a small retrieval corpus: they describe similarity, topicality, and detectable faults, and "
        "they do not establish factual correctness, empathy, or clinical appropriateness. Figure 4 "
        "summarizes these values.",
        section="4.3 Overall Performance",
        reason="All five values updated; interpretation restricted to what the metrics measure",
        issue="R07, R16, R17", source=SRC_OVR)

    # ==================================================================================
    # 4.4 BASELINE COMPARISON (Table 8)
    # ==================================================================================
    t8 = D.table_with("A0 \u2013 Generic LLM")
    t8_vals = {
        "A0 \u2013 Generic LLM": ["0.589", "0.081", "0.553", "\u2013 (no retrieval)", "1.41", "0.81"],
        "A1 \u2013 LLM + Standard RAG": ["0.613", "0.095", "0.581", "0.482", "0.98", "1.44"],
        "A2 \u2013 RAG + User Profile + Classifier Strategy": ["0.629", "0.104", "0.603", "0.615", "0.74", "2.21"],
        "A7 \u2013 Full PsyAdapt": ["0.634", "0.106", "0.609", "0.622", "0.69", "2.30"],
    }
    for key, vals in t8_vals.items():
        for j, v in enumerate(vals, start=1):
            D.upd_cell(t8, key, j, v, section="Table 8",
                       reason="Outdated value replaced with the verified updated result; grounding left "
                              "undefined for the configuration without retrieval",
                       issue="R07, R16", source=SRC_BASE)

    D.set_par(
        "From the Table 8 it can be concluded that the use of standard RAG for A1",
        "Table 8 shows that adding standard retrieval (A1) raises BERTScore-F1 from 0.589 to 0.613, "
        "ROUGE-L F1 from 0.081 to 0.095, and input relevance from 0.553 to 0.581 relative to the generic "
        "LLM baseline, while reducing detected defects from 1.41 to 0.98 per response. Adding user "
        "profiling and classifier-selected strategy guidance (A2) improves each of these further and "
        "raises grounding support from 0.482 to 0.615. Grounding support is not reported for A0 because "
        "that configuration retrieves no passages and the metric is undefined without them.",
        section="4.4 Baseline Comparison", reason="Values updated; grounding omission explained",
        issue="R07, R04", source=SRC_BASE)

    D.set_par(
        "The full version of the PsyAdapt architecture, A7, performs better in terms of",
        "The full configuration (A7) records the highest values on BERTScore-F1 (0.634), ROUGE-L F1 "
        "(0.106), input relevance (0.609), and grounding support (0.622), and the lowest defect count "
        "(0.69 per response), at the cost of the highest latency (2.3 s). The margin over A2 is small: "
        "+0.005 on BERTScore-F1, +0.002 on ROUGE-L F1, +0.006 on input relevance, and +0.007 on grounding "
        "support. Paired testing (Section 4.10) shows that the BERTScore-F1 differences from A0 and A1 "
        "remain significant after Holm correction, whereas the difference from A2 does not. The evidence "
        "in this table therefore supports the conclusion that retrieval and user profiling account for "
        "most of the measured improvement, and it does not support a claim that the full configuration "
        "outperforms A2.",
        section="4.4 Baseline Comparison",
        reason="'Performs better' claim replaced with the statistical outcome",
        issue="R02 Statistical - unsupported superiority claim", source=f"{SRC_BASE}, {SRC_STAT}")

    # ==================================================================================
    # 4.5 RESPONSE QUALITY ANALYSIS
    # ==================================================================================
    D.set_par(
        "Among the tested configurations, the proposed PsyAdapt system (A7) scored BERTScore-F1",
        "Among the tested configurations, the full PsyAdapt system (A7) recorded a BERTScore-F1 of 0.634 "
        "and a ROUGE-L F1 of 0.106, an input-relevance score of 0.609, and a grounding-support score of "
        "0.622. These are the highest values observed in the comparison, but the margin over "
        "configuration A2 is 0.005 on BERTScore-F1 and does not reach significance after correction for "
        "multiple comparisons (Section 4.10), so A7 is described here as numerically highest rather than "
        "as better.",
        section="4.5 Response Quality Analysis", reason="Values updated and claim strength matched to evidence",
        issue="R02, R07", source=f"{SRC_ABL}, {SRC_STAT}")

    D.set_par(
        "The score of the response defects is equal to 0.787",
        "The defect count for A7 is 0.69 per response, the lowest among the configurations, while its "
        "mean latency of 2.3 seconds is the highest, so the additional personalization operations buy a "
        "reduction in automatically detected faults at a cost in response time. Figure 5 shows the "
        "relationship between detected defects and latency across configurations. Figure 6 compares mean "
        "response defects across topics and configurations and indicates how defect levels vary by topic.",
        section="4.5 Response Quality Analysis", reason="Values updated; trade-off wording kept factual",
        issue="R07", source=SRC_ABL)

    D.append_par(
        "Figure 6. Defects by topic and configuration",
        "[PENDING \u2013 the topic-level breakdown plotted in this figure must be regenerated from the "
        "revised per-response evaluation output; the revised result set supplied for this revision "
        "contains the aggregate defect counts reported in Table 10 and Table 16 but not the topic-level "
        "breakdown.]",
        section="Figure 6 caption",
        reason="The plotted values predate the revision and the underlying per-topic data were not "
               "supplied, so the figure is flagged rather than silently retained",
        issue="R13 Figure", source="tables.zip (absent)", status="NOT RESOLVED")

    # ==================================================================================
    # 4.6 RESPONSE LENGTH
    # ==================================================================================
    D.set_par(
        "Figure 7 demonstrates that the generic LLM configuration (A0) yields relatively short responses",
        "Figure 7 shows that the generic LLM configuration (A0) yields the shortest responses, with a "
        "median of 128 words and an interquartile range of 96 to 171 words. The standard RAG "
        "configuration (A1) has a median of 141 words, and the personalized configurations are longer, "
        "with A7 at a median of 172 words and an interquartile range of 131 to 214 words. The full "
        "distribution statistics for every configuration are reported in Table 14. The suggested system "
        "therefore yields more detailed responses than the generic and the standard RAG configurations.",
        section="4.6 Response Length Analysis", reason="Median lengths updated to the verified values",
        issue="R18 Evaluation - length effect", source=SRC_LEN)

    D.append_par(
        "Overall, it is clear that PsyAdapt (A7) yields relatively detailed responses",
        "Section 4.12 quantifies this concern: across the evaluated responses, length correlates with "
        "grounding support (rho = 0.56) and with BERTScore-F1 (rho = 0.41), so part of the difference "
        "between A7 and the shorter configurations may be a length effect rather than a content effect.",
        section="4.6 Response Length Analysis",
        reason="Reviewer required that length be controlled for or reported alongside the metrics",
        issue="R18", source=SRC_LCORR)
    return D



def apply_section4b(D):
    # ==================================================================================
    # 4.7 SOTA COMPARISON (Table 9 rebuilt, DESIGN A)
    # ==================================================================================
    D.set_par(
        "PsyAdapt A7 was compared with three open weight large language models",
        "PsyAdapt was compared with three open-weight large language models: Gemma 3 4B, Llama 3.2 3B "
        "Instruct, and Mistral 7B Instruct. The comparison reported in the previous version of this paper "
        "has been withdrawn for two reasons: it compared Qwen2.5-3B inside the framework against the "
        "other models outside it, which confounds the generator with the pipeline, and it reported a "
        "grounding-support value for models that performed no retrieval, although grounding support is "
        "defined only against retrieved passages. The revised comparison uses a generator-swap design in "
        "which each model is placed inside the identical pipeline and evaluated on the same 300 benchmark "
        "items with the same prompts, the same retrieval cache, the same greedy decoding settings, and the "
        "same metric implementations, so that the only difference between arms is the generator. Two arms "
        "are run per model: a bare arm with no retrieval, profile, or strategy guidance, which is the "
        "counterpart of A0, and a full-context arm, which is the counterpart of A7. Grounding support is "
        "computed only for the arms that actually retrieved passages and is left undefined for the bare "
        "arms.",
        section="4.7 SOTA Model Comparison",
        reason="Comparison redesigned as a generator swap inside the same pipeline; invalid grounding "
               "values for bare models removed",
        issue="R03, R04, R05 - constructed and confounded SOTA table; R31 required action",
        source=f"{SRC_SOTA}, {SRC_NB}")

    D.set_par(
        "Table 9. Comparative Evaluation of PsyAdapt with Selected LLM Baselines",
        "Table 9. Comparative Evaluation of PsyAdapt with Selected Open-Weight LLMs under the "
        "Generator-Swap Design (identical benchmark, prompts, retrieval, decoding, and metrics)",
        section="Table 9 caption", reason="Caption reflects the revised comparison design",
        issue="R03, R05", source=SRC_SOTA)

    t9 = D.table_with("Input relevance \u2191")
    D.replace_table(t9, T9_HEADER, T9_ROWS, widths=[2400, 1400, 900, 1000, 1100, 1100, 1126], section="Table 9",
                    reason="Previous values were unverifiable and monotone, and assigned grounding scores "
                           "to models without retrieval; replaced by measured generator-swap results with "
                           "bootstrap confidence intervals",
                    issue="R03, R04", source=SRC_SOTA)

    D.set_par(
        "As presented in Table 9, PsyAdapt A7 obtained the best BERTScore-F1",
        "Table 9 and Figure 8 report the results. Among the full-context arms, the Qwen-based "
        "configuration records the highest mean BERTScore-F1 (0.634, 95% CI [0.628, 0.640]), followed by "
        "Mistral 7B Instruct (0.630, [0.624, 0.636]), Gemma 3 4B (0.627, [0.621, 0.633]), and Llama 3.2 "
        "3B Instruct (0.623, [0.617, 0.629]). These intervals overlap, so the ordering among generators is "
        "not established by this experiment and the differences are numerical only. The consistent "
        "separation is between context conditions rather than between models: every full-context arm "
        "exceeds its own bare counterpart by roughly 0.04 to 0.05 BERTScore-F1 and records roughly half as "
        "many detected defects per response, and the bare Qwen configuration (A0: BERTScore-F1 0.589, 1.41 "
        "defects) lies in the same range as the other bare models (0.580 to 0.586, 1.55 to 1.70 defects). "
        "The comparison therefore supports the claim that the PsyAdapt context improves the measured "
        "metrics for every generator tested, and it does not support a claim that a particular base model "
        "is superior or that the framework's advantage depends on the choice of generator.",
        section="4.7 SOTA Model Comparison",
        reason="Superiority claims replaced by the confidence-interval reading of the revised results",
        issue="R03, R05, R02", source=SRC_SOTA)

    D.insert_block(
        "The comparison therefore supports the claim that the PsyAdapt context improves",
        [("FIG:" + FIGDIR + "/figure8_sota_design_a.png", ""),
         ("figcap",
          "Figure 8. Generator-swap comparison inside the identical PsyAdapt pipeline: BERTScore-F1 mean "
          "with 95% bootstrap confidence interval for each bare and full-context arm.")],
        section="4.7 SOTA Model Comparison (new Figure 8)",
        reason="Reviewer required a defensible comparison; the confidence intervals are shown so that the "
               "overlap between generators is visible",
        issue="R03, R05, R32", source=SRC_SOTA)

    # ==================================================================================
    # 4.8 ABLATION STUDY (Table 10)
    # ==================================================================================
    t10 = D.table_with("A3 \u2013 Without Emotion Detection")
    t10_vals = {
        "A2 \u2013 RAG + User Profile + Classifier Strategy": ["0.629", "0.104", "0.603", "0.615", "0.74", "2.21"],
        "A3 \u2013 Without Emotion Detection": ["0.621", "0.099", "0.592", "0.598", "0.86", "2.12"],
        "A4 \u2013 Without Personality Detection": ["0.626", "0.102", "0.601", "0.608", "0.76", "2.19"],
        "A5 \u2013 Without Memory": ["0.629", "0.104", "0.603", "0.614", "0.74", "2.20"],
        "A6 \u2013 Without Strategy Guidance": ["0.617", "0.097", "0.593", "0.604", "0.89", "1.87"],
        "A7 \u2013 Full PsyAdapt": ["0.634", "0.106", "0.609", "0.622", "0.69", "2.30"],
    }
    for key, vals in t10_vals.items():
        for j, v in enumerate(vals, start=1):
            D.upd_cell(t10, key, j, v, section="Table 10",
                       reason="Outdated value replaced with the verified updated result",
                       issue="R02, R07", source=SRC_ABL)
    D.upd_cell(t10, "A5 \u2013 Without Memory", 0, "A5 \u2013 Without Memory (single-turn control)",
               section="Table 10",
               reason="Row relabelled because a single-turn benchmark cannot exercise memory",
               issue="R13", source=SRC_ABL)

    D.set_par(
        "As can be seen from the Table 10 results, A7 configuration obtains the best BERTScore-F1",
        "Table 10 shows that the full configuration records the highest BERTScore-F1 (0.634), ROUGE-L F1 "
        "(0.106), input relevance (0.609), and grounding support (0.622) among the ablation "
        "configurations, and the lowest defect count (0.69 per response). Removing emotion detection (A3) "
        "lowers BERTScore-F1 to 0.621 and grounding support to 0.598 and raises defects to 0.86; removing "
        "strategy guidance (A6) lowers BERTScore-F1 to 0.617 and raises defects to 0.89, the highest value "
        "among the personalized configurations. Both differences from A7 remain significant after Holm "
        "correction (Table 12), so emotion conditioning and strategy guidance do affect the measured "
        "outcomes, with their clearest effect on the number of automatically detected defects.",
        section="4.8 Ablation Study", reason="Values updated and claims tied to the corrected statistics",
        issue="R02, R07", source=f"{SRC_ABL}, {SRC_STAT}")

    D.set_par(
        "A4 configuration obtains almost identical results compared with A7",
        "Removing personality detection (A4) produces a smaller change: BERTScore-F1 falls to 0.626, input "
        "relevance to 0.601, and grounding support to 0.608, while defects rise to 0.76. The paired "
        "difference from A7 (+0.008 BERTScore-F1, Holm-corrected p = 0.019) is significant but is the "
        "smallest of the component ablations, which is consistent with the weak accuracy of the "
        "personality classifier reported in Table 11 and with the fact that the personality signal is "
        "injected only when its confidence exceeds the validation-selected threshold.",
        section="4.8 Ablation Study", reason="Values updated; effect size interpreted against classifier quality",
        issue="R02, R09", source=f"{SRC_ABL}, {SRC_STAT}, {SRC_CLS}")

    D.set_par(
        "A5 configuration, without memory, obtains the same values of BERTScore-F1",
        "The A5 configuration differs from A2 only in that the conversation-memory element is switched "
        "off, and on a single-turn benchmark there is no conversation history to insert, so the two "
        "configurations receive effectively the same prompt: their BERTScore-F1, ROUGE-L F1, input "
        "relevance, and defect values are identical, and grounding support differs by 0.001. This row is "
        "therefore a control confirming that the ablation is inert in a single-turn setting, and it "
        "carries no information about the value of memory; the memory component is evaluated separately on "
        "multi-turn conversations in Section 4.11. Configuration A2 and configuration A6 differ only in "
        "strategy guidance: A2 inserts the strategy predicted by the classifier into the prompt, whereas "
        "A6 omits the strategy-guidance block entirely, and A7 replaces the classifier prediction with the "
        "adaptive strategy scorer described in Section 3.7.",
        section="4.8 Ablation Study",
        reason="Memory row reinterpreted as a control and the previously unstated A2/A6 difference explained",
        issue="R13, R22", source=f"{SRC_ABL}, {SRC_NB}")

    D.set_par(
        "Concerning the latency, A7 configuration has the highest latency with 2.395 seconds",
        "Latency follows the amount of context and processing: among the personalized configurations A6 is "
        "fastest at 1.87 seconds because it omits the strategy-guidance block, and A7 is slowest at 2.3 "
        "seconds. The ablation values are single-run point estimates, and they are accompanied by the "
        "paired analysis in Section 4.10; differences on metrics for which no paired test is reported "
        "should be treated as descriptive only.",
        section="4.8 Ablation Study", reason="Latency values updated and the statistical caveat retained",
        issue="R02, R07", source=SRC_ABL)

    # ==================================================================================
    # NEW SUBSECTIONS 4.9 - 4.16
    # ==================================================================================
    anchor = D.find("Latency follows the amount of context and processing")._p

    blocks = [
        # ---------------- 4.9 ----------------
        ("4.9 (new subsection)", "Component-level classifier accuracy was never reported",
         "R09, R10, R33", f"{SRC_CLS}, {SRC_NB}", "RESOLVED", [
             ("head", "4.9 Component-Level Classifier Performance"),
             ("body",
              "The personalization signals used by the framework are produced by trained components, and "
              "their accuracy bounds how informative those signals can be. Each component was evaluated "
              "on the held-out test split of its own dataset, and the results are reported in Table 11. "
              "The per-response predictions from which per-class confusion matrices are derived are "
              "exported by the evaluation pipeline."),
             ("cap", "Table 11. Component-Level Classifier Performance on Held-Out Test Splits"),
             ("TABLE", (T11_HEADER, T11_ROWS, [1900, 1800, 1500, 1000, 950, 950, 926])),
             ("body",
              "The emotion classifier reaches 0.78 accuracy and 0.72 macro-F1, the strongest of the three "
              "components, which supports its use as a cautious prompt signal. The personality component "
              "reaches 0.52 accuracy and 0.44 macro-F1 for the primary Big Five trait predicted from a "
              "single user message. This is weak, and it is the reason the personality signal is gated by "
              "a confidence threshold, is phrased in the prompt as a tentative communication preference, "
              "and is never presented to the user as a description of their personality. The deployed "
              "strategy classifier, which sees only the dialogue context preceding the supporter turn, "
              "reaches 0.61 accuracy and 0.53 macro-F1, so its output is used as soft guidance rather "
              "than as a decision."),
             ("body",
              "The last row of Table 11 is a diagnostic. An earlier version of the strategy classifier "
              "received the supporter response as part of its input and reached 0.87 accuracy; because "
              "the supporter response is the text whose strategy is being predicted, that figure reflects "
              "label leakage rather than predictive ability at inference time. The leaky variant is "
              "reported only to document the problem and is not used in any configuration or result in "
              "this paper."),
         ]),
        # ---------------- 4.10 ----------------
        ("4.10 (new subsection)",
         "Reviewer required bootstrap confidence intervals and paired tests for every configuration "
         "comparison, and that every 'better' claim be rewritten to match",
         "R02, R32", f"{SRC_STAT}, {SRC_NB}", "RESOLVED", [
             ("head", "4.10 Statistical Analysis of Configuration Differences"),
             ("body",
              "A single evaluation run produces point estimates only, so the configuration differences "
              "reported in Table 8 and Table 10 were re-analysed as paired comparisons over the same 300 "
              "benchmark items. For each contrast, per-item differences in BERTScore-F1 were tested with "
              "a two-sided Wilcoxon signed-rank test, summarized with a 95% bootstrap confidence interval "
              "computed from 5,000 resamples and with Cohen's d_z, and corrected for multiple comparisons "
              "using the Holm procedure. Table 12 and Figure 9 report the outcome."),
             ("cap", "Table 12. Paired Contrasts against the Full PsyAdapt Configuration "
                     "(BERTScore-F1, n = 300 benchmark items)"),
             ("TABLE", (T12_HEADER, T12_ROWS, [1100, 1200, 1500, 1000, 1100, 1300, 1826])),
             ("body",
              "Three readings follow. First, the full configuration is separated from the two weaker "
              "baselines by margins that survive correction: +0.045 against the generic LLM (95% CI "
              "[0.031, 0.058], d_z = 0.61, Holm-corrected p = 0.0009) and +0.021 against standard RAG "
              "(95% CI [0.011, 0.031], d_z = 0.42, Holm-corrected p = 0.0016). Second, the differences "
              "from the emotion, personality, and strategy-guidance ablations are small but consistent "
              "and also survive correction: +0.013 against A3 (Holm-corrected p = 0.0028), +0.008 against "
              "A4 (p = 0.0186), and +0.017 against A6 (p = 0.0021). Third, the difference from A2, the "
              "configuration that combines retrieval with user profiling and classifier-selected "
              "strategy, is +0.005 with a confidence interval that reaches zero ([0.000, 0.010]), d_z = "
              "0.16, and an uncorrected p-value of 0.041 that becomes 0.287 after Holm correction; the "
              "same holds for the single-turn memory control A5. The full configuration therefore "
              "obtained a numerically higher score than A2 and A5, although the difference was not "
              "statistically significant, and no claim of superiority over A2 is made in this paper."),
             ("FIG:" + FIGDIR + "/figure9_paired_contrasts.png", ""),
             ("figcap",
              "Figure 9. Paired mean differences in BERTScore-F1 between the full PsyAdapt configuration "
              "and each other configuration, with 95% bootstrap confidence intervals; filled markers "
              "indicate contrasts that remain significant after Holm correction."),
             ("body",
              "This analysis also fixes the interpretation of the remaining metrics. Paired contrasts are "
              "currently available for BERTScore-F1, the primary automatic metric; differences in ROUGE-L "
              "F1, input relevance, grounding support, defects, and latency are reported descriptively, "
              "and wording such as higher or lower in those cases refers to observed means only. "
              "[PENDING: paired tests for the remaining automatic metrics are produced by the same "
              "procedure in the revision pipeline and should be inserted here when the corresponding "
              "result file is available.]"),
         ]),
        # ---------------- 4.11 ----------------
        ("4.11 (new subsection)",
         "Reviewer required that memory be removed or evaluated on multi-turn data",
         "R13, R36", f"{SRC_MEM}, {SRC_NB}", "RESOLVED", [
             ("head", "4.11 Multi-Turn Memory Experiment"),
             ("body",
              "The conversation-memory component cannot be evaluated on the single-turn benchmark, "
              "because a single-turn item contains no earlier turns to remember; the A5 row of Table 10 "
              "is a control rather than evidence about memory. Memory was therefore evaluated in a "
              "separate paired experiment on multi-turn conversations from the ESConv test set. Forty "
              "conversations with at least four user turns were sampled with a fixed seed. For each "
              "conversation the system generates a reply to the fourth user turn under two conditions "
              "that are identical in retrieval, user profile, strategy selection, and decoding settings "
              "and differ only in whether the dynamic user model's memory of the first three user turns "
              "is inserted into the prompt: the with-memory arm (A2-M) includes it, and the without-memory "
              "arm (A5-M) does not."),
             ("cap", "Table 13. Multi-Turn Memory Experiment on ESConv Conversations"),
             ("TABLE", (T13_HEADER, T13_ROWS, [2100, 1200, 1200, 1100, 1300, 1000, 1126])),
             ("body",
              "With memory, mean BERTScore-F1 against the reference reply is 0.618, compared with 0.584 "
              "without memory; the paired difference is +0.034 (95% CI [0.001, 0.066]), d_z = 0.34, "
              "Wilcoxon p = 0.048. The direction favours memory, but the sample comprises 40 "
              "conversations, the confidence interval nearly reaches zero, and the comparison uses a "
              "single automatic metric against one reference reply. The experiment is therefore reported "
              "as preliminary evidence that conversation memory can change multi-turn replies, and not as "
              "a demonstration that the memory component improves support quality."),
         ]),
        # ---------------- 4.12 ----------------
        ("4.12 (new subsection)",
         "Reviewer required that response length be controlled for or reported alongside the metrics",
         "R18", f"{SRC_LEN}, {SRC_LCORR}", "PARTIALLY RESOLVED", [
             ("head", "4.12 Response Length and Metric Validity"),
             ("body",
              "Response length differs systematically across configurations, and several of the automatic "
              "metrics are sensitive to it, so length is reported alongside the metrics rather than "
              "treated as neutral. Table 14 gives the length distribution per configuration in words. All "
              "configurations were generated under the same instruction to aim at approximately 150 to "
              "200 words, so the differences arise from the amount of context supplied rather than from "
              "different length targets."),
             ("cap", "Table 14. Response-Length Distribution across PsyAdapt Configurations (words)"),
             ("TABLE", (T14_HEADER, T14_ROWS, [4226, 1400, 1100, 1100, 1200])),
             ("cap", "Table 15. Spearman Correlations between Response Length and the Automatic Metrics"),
             ("TABLE", (T15_HEADER, T15_ROWS, [3026, 3500, 2500])),
             ("body",
              "Table 15 reports Spearman correlations between response length and each metric across the "
              "evaluated responses. Grounding support is the most length-dependent metric (rho = 0.56, p "
              "< 0.001), followed by BERTScore-F1 (rho = 0.41, p = 0.0001) and ROUGE-L F1 (rho = 0.29, p "
              "= 0.003), while input relevance is only weakly related to length (rho = 0.18, p = 0.031). "
              "Because the full configuration also produces the longest responses, part of its advantage "
              "on BERTScore-F1 and grounding support may reflect length rather than better content. This "
              "is a direct caveat on Table 7, Table 8, and Table 10, and it reinforces the conclusion of "
              "Section 4.10 that the small differences among the personalized configurations should not "
              "be read as demonstrated quality gains. [PENDING: the revision pipeline also recomputes "
              "every contrast within matched response-length strata; that length-controlled comparison "
              "should be inserted here once the corresponding result file is available.]"),
         ]),
        # ---------------- 4.13 ----------------
        ("4.13 (new subsection)",
         "Reviewer required that the defect detector be defined and that safety screening be described",
         "R12, R07, R35", f"{SRC_DEF}, {SRC_NB}", "RESOLVED", [
             ("head", "4.13 Safety Screening and Defect Composition"),
             ("body",
              "The screens described in Section 3.9.2 and Section 3.10 were applied to every generated "
              "response, and the defect counts reported throughout Section 4 are their output. Table 16 "
              "reports the composition of the 0.69 defects per response recorded for the full "
              "configuration."),
             ("cap", "Table 16. Defect Composition of the Full PsyAdapt Configuration (A7)"),
             ("TABLE", (T16_HEADER, T16_ROWS, [5526, 3500])),
             ("body",
              "The largest category is unsupported medical claims (0.21 per response), followed by "
              "overstated certainty (0.14) and prescriptive advice without a disclaimer (0.12). Missing "
              "referral on inputs that the input screen flagged as high distress occurs at 0.09 per "
              "response; although it is only the fourth largest category, it carries the most direct "
              "safety implication, because it means that a non-negligible share of flagged inputs "
              "received a response containing no referral expression. Diagnostic overreach (0.08) and "
              "inappropriate boundary statements (0.05) are the least frequent categories."),
             ("body",
              "Two limits on the interpretation of these counts must be stated. First, the detection "
              "procedure is automated: the rule-based component matches surface patterns and can miss "
              "paraphrases or flag acceptable phrasing, and the annotation component is an "
              "instruction-tuned model, not a clinician. Second, a lower count is evidence of fewer "
              "detected faults of these specific kinds, and not evidence that a response is safe, "
              "empathetic, or clinically appropriate. No claim of clinical safety is made anywhere in "
              "this paper on the basis of these counts."),
         ]),
        # ---------------- 4.14 ----------------
        ("4.14 (new subsection)",
         "Reviewer required a blinded human evaluation on 50-100 items; the protocol is implemented but "
         "no annotations exist, so no human results are reported and the gap is stated",
         "R07, R34", SRC_NB, "REQUIRES HUMAN ACTION", [
             ("head", "4.14 Human Evaluation: Protocol and Current Status"),
             ("body",
              "The automatic metrics used in Section 4.3 to Section 4.8 cannot determine whether a "
              "response is empathetic, helpful, or safe for a distressed user, so a blinded human "
              "evaluation is part of the evaluation design and its protocol is implemented in the "
              "pipeline. For a fixed sample of benchmark items, the responses of two configurations are "
              "presented side by side as reply A and reply B with the presentation order randomized; the "
              "configuration identities are held in a separate key that annotators never receive; "
              "annotators select a winner or a tie for each criterion; and the written guidelines state "
              "that length, polish, and formatting are not merits. The contrasts exported for annotation "
              "are the full configuration against A2, A2 against standard RAG, and A2 against the "
              "configuration without strategy guidance."),
             ("body",
              "No annotated returns are available at the time of this revision. Consequently this paper "
              "reports no human-evaluation results and claims none: every quantitative result in Section "
              "4 is automatic. Completing the annotation, reporting inter-annotator agreement and "
              "agreement between human and automatic judgments, and obtaining ratings from annotators "
              "with clinical training remain required work before any claim about empathy, helpfulness, "
              "or safety can be made. [REQUIRES HUMAN ACTION]"),
             ("body",
              "In place of human ratings, the pipeline implements automated judging protocols: a blinded "
              "pairwise forced choice in which each pair is judged in both presentation orders so that "
              "position bias is measured rather than assumed, a structured defect-counting protocol, and "
              "a pointwise replication that documents the rating saturation observed previously. Of "
              "these, only the defect counts are reported in this paper, in Section 4.13. Automated "
              "judging by a language model is not a substitute for human or clinical evaluation."),
         ]),
        # ---------------- 4.15 ----------------
        ("4.15 (new subsection)",
         "Reviewer noted there is no Discussion section",
         "R25", f"{SRC_ABL}, {SRC_STAT}, {SRC_SOTA}", "RESOLVED", [
             ("head", "4.15 Discussion"),
             ("body",
              "Read together, the results support a narrower claim than the configuration names suggest. "
              "The largest and most reliable gains come from the two simplest additions: retrieval, which "
              "raises BERTScore-F1 from 0.589 to 0.613 and reduces detected defects from 1.41 to 0.98 per "
              "response relative to the generic model, and user profiling with classifier-selected "
              "strategy guidance, which adds a further 0.016 on BERTScore-F1 and raises grounding support "
              "from 0.482 to 0.615. What the framework adds beyond that point is small on the "
              "reference-similarity metrics and is not statistically separable from A2 (Section 4.10)."),
             ("body",
              "The clearer contribution of the remaining components is on defects rather than on "
              "similarity. Removing strategy guidance raises defects from 0.69 to 0.89 per response and "
              "removing emotion detection raises them to 0.86, whereas removing personality detection "
              "changes them comparatively little (0.76). The adaptive strategy scorer used in the full "
              "configuration also records fewer defects than the classifier-selected strategy of A2 (0.69 "
              "against 0.74). Since these counts come from automated detection, the appropriate reading "
              "is that emotion-conditioned and strategy-conditioned prompting reduces the rate of "
              "detectable faults of specific kinds, which is a plausible mechanism, and not that it "
              "produces better support."),
             ("body",
              "The generator-swap comparison locates the effect in the pipeline rather than in the base "
              "model: the three external generators inside PsyAdapt fall within overlapping confidence "
              "intervals of one another and of the Qwen-based configuration, while each of them, and Qwen "
              "itself, is clearly separated from its own bare counterpart. This is consistent with the "
              "ablation evidence, although the present study evaluates a single dialogue dataset, a "
              "single decoding setting, and one run per arm."),
             ("body",
              "Finally, the metric analysis limits how these numbers can be used at all. Reference-based "
              "similarity is computed against one of many acceptable replies; grounding support measures "
              "textual overlap with retrieved passages rather than factual correctness; and both "
              "correlate with response length (Section 4.12). The framework's measurable effects are "
              "therefore established on proxies. Whether they correspond to support that users experience "
              "as helpful, and whether the system is safe for distressed users, is not answered by this "
              "evaluation and requires the human and clinical studies described in Section 4.14."),
         ]),
        # ---------------- 4.16 ----------------
        ("4.16 (new subsection)",
         "Reviewer noted the absence of a Limitations section",
         "R25, R07, R14, R17, R18", f"{SRC_ABL}, {SRC_CLS}, {SRC_NB}", "RESOLVED", [
             ("head", "4.16 Limitations"),
             ("body",
              "The following limitations qualify every result reported above. Evaluation is automatic "
              "only: BERTScore-F1, ROUGE-L F1, input relevance, grounding support, and the defect counts "
              "are proxies computed by software, and none of them measures empathy, helpfulness, clinical "
              "appropriateness, or harm. No clinical or expert validation was performed, and no human "
              "ratings were available for this revision (Section 4.14)."),
             ("body",
              "The reference-based metrics compare each response with a single MentalChat16K reply, "
              "although many different replies are acceptable, and ROUGE-L F1 stays below 0.11 for every "
              "configuration, so lexical overlap discriminates weakly. Response length is a confound: it "
              "correlates with grounding support and BERTScore-F1, and the length-controlled comparison "
              "is not yet available (Section 4.12)."),
             ("body",
              "Component quality limits the personalization. The personality classifier reaches 0.52 "
              "accuracy for the primary trait from a single message, the deployed strategy classifier "
              "reaches 0.61 accuracy, and the need/context module is rule-based and has not been "
              "evaluated against labelled annotations, so some personalization signals may be noisy. The "
              "psychoeducational knowledge base is small, which limits what retrieval can supply and "
              "means grounding is measured against a narrow set of passages."),
             ("body",
              "The main benchmark is single-turn, so the evaluation does not test sustained dialogue, and "
              "the multi-turn memory evidence rests on 40 conversations and one automatic metric. Results "
              "come from one generation run with greedy decoding: the paired tests address item-level "
              "variation, not run-to-run variation, and paired testing is currently reported for "
              "BERTScore-F1 only. Latency was measured at batch size 1 in a single GPU session and is "
              "hardware-specific. Finally, the study covers English-language text, one dialogue dataset, "
              "and generators in the 3B to 7B range, so the findings should not be generalized beyond "
              "that setting."),
         ]),
    ]

    for section, reason, issue, source, status, items in blocks:
        anchor = D.insert_block(None, items, after_element=anchor, section=section,
                                reason=reason, issue=issue, source=source, status=status)

    # ==================================================================================
    # 5. CONCLUSION
    # ==================================================================================
    D.set_par(
        "Our full PsyAdapt model (A7) attained BERTScore-F1 = 0.632",
        "Our full PsyAdapt configuration (A7) attained BERTScore-F1 = 0.634, ROUGE-L F1 = 0.106, input "
        "relevance = 0.609, grounding support = 0.622, 0.69 defects per response, and 2.3 s latency on the "
        "fixed 300-item benchmark. Paired analysis shows that this configuration is separated from the "
        "generic-LLM and standard-RAG baselines by margins that survive Holm correction, and from the "
        "emotion, personality, and strategy-guidance ablations by smaller but corrected-significant "
        "margins, whereas its advantage over the retrieval-plus-profiling configuration A2 (+0.005 "
        "BERTScore-F1, 95% CI [0.000, 0.010], Holm-corrected p = 0.287) is not statistically significant. "
        "In the generator-swap comparison, in which Gemma 3 4B, Llama 3.2 3B Instruct, and Mistral 7B "
        "Instruct were evaluated inside the identical pipeline, all full-context arms fall within "
        "overlapping confidence intervals, while every full-context arm is clearly separated from its own "
        "bare counterpart. The ablation study indicates that retrieval and user profiling account for most "
        "of the measured gain, and that emotion detection and adaptive strategy guidance mainly reduce the "
        "number of automatically detected response defects. A separate multi-turn experiment on 40 ESConv "
        "conversations provides preliminary evidence that conversation memory affects replies in "
        "multi-turn settings; the single-turn ablation does not test memory. No human or clinical "
        "evaluation was available for this study, so no claim about empathy, helpfulness, clinical safety, "
        "or therapeutic effect is made.",
        section="5. Conclusion and Future Work",
        reason="Conclusion aligned with the verified results; claim of superiority over three stand-alone "
               "LLMs and over A2 withdrawn",
        issue="R02, R03, R07, R13", source=f"{SRC_OVR}, {SRC_STAT}, {SRC_SOTA}, {SRC_MEM}")

    D.append_par(
        "Future work will include the development of evaluation on more datasets",
        "Specifically, the blinded pairwise human-evaluation protocol described in Section 4.14 will be "
        "completed with several annotators, including annotators with clinical training, and reported with "
        "inter-annotator agreement; paired significance testing will be extended to every automatic metric "
        "and to length-matched strata; the psychoeducational knowledge base will be enlarged; and the "
        "memory component will be evaluated on a larger multi-turn benchmark.",
        section="5. Conclusion and Future Work",
        reason="Future work made specific to the outstanding reviewer requirements",
        issue="R34, R36, R14", source=SRC_NB)
    return D


def main():
    for hl, out in [(True, f"{OUTDIR}/UPDATED_REFERENCE_PAPER.docx"),
                    (False, f"{OUTDIR}/UPDATED_CLEAN_PAPER.docx")]:
        CHANGELOG.clear()
        _seen_log.clear()
        D = Doc(SRC, highlight=hl)
        apply_edits(D)
        apply_section4(D)
        apply_section4b(D)
        apply_addendum(D)
        D.save(out)
        replace_media(out, {
            "word/media/image4.png": f"{FIGDIR}/figure4_overall_performance_A7.png",
            "word/media/image5.png": f"{FIGDIR}/figure5_quality_vs_latency.png",
            "word/media/image7.png": f"{FIGDIR}/figure7_response_lengths.png",
        })
        print("wrote", out, "| changes logged:", len(CHANGELOG))

    with open(f"{OUTDIR}/CHANGE_LOG.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Section", "Original Text/Value", "Updated Text/Value", "Reason",
                    "Reviewer Issue", "Source", "Status"])
        w.writerows(CHANGELOG)
    print("wrote CHANGE_LOG.csv with", len(CHANGELOG), "rows")





# ======================================================================================
# ADDENDUM: Table 2, Table 3, Algorithm 1
# ======================================================================================
def _algorithm_cell(D):
    for t in D.d.tables:
        if len(t.rows) == 1 and "Input: MentalChat16K" in t.rows[0].cells[0].text:
            return t.rows[0].cells[0]
    raise KeyError("algorithm cell not found")


def _rewrite_cell_par(D, p, new_text):
    rpr = D._first_rpr(p)
    for child in list(p._p):
        if child.tag != qn("w:pPr"):
            p._p.remove(child)
    D._add_run(p, new_text, rpr, True)


def apply_addendum(D):
    # ---- Table 2: knowledge-base scale ----
    t2 = D.table_with("NIH / NHLBI, WHO, NIMH resources")
    D.upd_cell(t2, "NIH / NHLBI, WHO, NIMH resources", 3,
               "81 coarse chunks, re-chunked into ~500-character fine passages for retrieval "
               "[UNVERIFIED chunk counts]",
               section="Table 2",
               reason="Reflects the re-chunking used for retrieval; counts flagged for author confirmation",
               issue="R14 Dataset - knowledge base size", source=SRC_NB, status="PARTIALLY RESOLVED")

    # ---- Table 3: feature representations ----
    t3 = D.table_with("Strategy data")
    D.upd_cell(t3, "Personality data", 3, "TF-IDF feature representation (linear SVM)",
               section="Table 3", reason="Representation corrected to the implemented component",
               issue="R09", source=SRC_CLS)
    D.upd_cell(t3, "Strategy data", 2,
               "Extraction of the dialogue history preceding the supporter turn (no supporter response); "
               "strategy-label preparation",
               section="Table 3", reason="Leakage-free input construction documented",
               issue="R09", source=SRC_NB)
    D.upd_cell(t3, "Strategy data", 3, "TF-IDF (1-2 gram) feature representation (logistic regression)",
               section="Table 3", reason="Representation corrected to the deployed classifier",
               issue="R09", source=SRC_CLS)

    # ---- Algorithm 1 ----
    cell = _algorithm_cell(D)
    algo_edits = {
        "Divide the collected knowledge documents into smaller text chunks.":
            "Divide the collected knowledge documents into text chunks of approximately 500 characters "
            "with 80-character overlap at sentence boundaries.",
        "Train the DistilBERT-based support-strategy classifier using ESConv.":
            "Train the context-only support-strategy classifier (TF-IDF features over the dialogue "
            "history preceding the supporter turn) using ESConv.",
        "Receive the user's input .": "Receive the user's input X.",
        "Predict the user's emotional state using the emotion classifier.":
            "Predict the user's emotional state E using the emotion classifier.",
        "Estimate the user's personality profile using the personality classifier.":
            "Estimate the user's personality profile P using the personality classifier.",
        "Identify the user's need and contextual information .":
            "Identify the user's need and contextual information N using the rule-based need/context module.",
        "Construct the dynamic user representation using .":
            "Construct the dynamic user representation U = (E, P, N).",
        "Select an adaptive support strategy based on the user representation.":
            "Select an adaptive support strategy S based on the user representation.",
        "Retrieve relevant psychoeducational knowledge using the user input and contextual information.":
            "Retrieve relevant psychoeducational knowledge R using the user input and contextual information.",
        "Construct the personalized generation context using .":
            "Construct the personalized generation context C = (X, U, S, R).",
        "Generate a personalized psychoeducational response .":
            "Generate a personalized psychoeducational response Y.",
        "Apply safety and quality verification to assess response relevance, grounding, and potential defects.":
            "Apply the input safety screen, the safety system prompt, and the response screen to assess "
            "response relevance, grounding, and potential defects.",
        "Return the verified response to the user.": "Return the verified response Y* to the user.",
        "Evaluate the framework using BERTScore-F1, ROUGE-L F1, input relevance, grounding support, "
        "defects per response, and latency.":
            "Evaluate the framework using BERTScore-F1, ROUGE-L F1, input relevance, grounding support, "
            "defects per response, and latency, with paired Wilcoxon tests, bootstrap confidence "
            "intervals, and Holm correction.",
        "Perform baseline comparisons and ablation experiments to assess the contribution of individual "
        "components.":
            "Perform baseline comparisons, ablation experiments, a generator-swap comparison, and a "
            "separate multi-turn memory experiment to assess the contribution of individual components.",
    }
    for p in cell.paragraphs:
        txt = p.text.strip()
        if txt in algo_edits:
            _rewrite_cell_par(D, p, algo_edits[txt])
            log("Algorithm 1", txt, algo_edits[txt],
                "Step corrected to the implemented procedure / equation symbols restored as plain text",
                "R09, R13, R23, R02", SRC_NB)

    # insert the threshold-selection step after the strategy-classifier step
    for p in cell.paragraphs:
        if p.text.strip().startswith("Train the context-only support-strategy classifier"):
            new_tr = deepcopy(p._p)
            p._p.addnext(new_tr)
            newp = Paragraph(new_tr, p._parent)
            for child in list(new_tr):
                if child.tag != qn("w:pPr"):
                    new_tr.remove(child)
            D._add_run(newp,
                       "Select the emotion and personality injection thresholds on the MentalChat16K "
                       "validation split (never on the benchmark).",
                       D._first_rpr(p) or D.cell_rpr, True)
            log("Algorithm 1", "(new step)",
                "Select the emotion and personality injection thresholds on the validation split.",
                "Documents the threshold-selection stage the reviewer found unexplained",
                "R20", SRC_NB)
            break
    return D


if __name__ == "__main__":
    main()
