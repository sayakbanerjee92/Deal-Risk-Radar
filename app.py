"""Deal Risk Radar — deterministic, local, no-API contract screening."""

from __future__ import annotations

import io
import re
from dataclasses import asdict, dataclass
from typing import Literal

import streamlit as st
from docx import Document
from pypdf import PdfReader

APP_NAME = "Deal Risk Radar"
DISCLAIMER = (
    "This is a deterministic first-pass commercial diligence tool, not legal advice. "
    "Qualified counsel must verify material conclusions before a deal or signing decision."
)

CATEGORY_PATTERNS = {
    "Change of Control / Assignment": r"\b(change of control|assignment|assign(?:ment)?|successors? and assigns?)\b",
    "Termination / Renewal": r"\b(termination|terminate|renewal|auto[- ]renew)\b",
    "Revenue / Payment": r"\b(fees?|payment|net\s+\d+|invoice|minimum commitment|committed spend)\b",
    "Pricing / MFN / Exclusivity": r"\b(most favou?red|MFN|exclusiv|rebate|price cap|pricing)\b",
    "Limitation of Liability": r"\b(liability|damages|liability cap|uncapped|unlimited)\b",
    "Indemnity": r"\b(indemnif|defend|hold harmless)\b",
    "Intellectual Property": r"\b(intellectual property|background IP|foreground IP|deliverable|work product)\b",
    "Data Protection / Security": r"\b(data protection|personal data|security|data breach|incident)\b",
    "Audit / Compliance": r"\b(audit|compliance|inspection)\b",
    "Service Levels / Credits": r"\b(service level|SLA|service credit|availability)\b",
    "Restrictive Covenants": r"\b(non[- ]solicit|non[- ]compete|publicity|hire)\b",
    "Dispute Resolution / Governing Law": r"\b(governing law|jurisdiction|arbitration|dispute)\b",
}

RULES = [
    ("Change of Control / Assignment", "RED", r"\bchange of control\b.{0,180}\b(terminate|termination|consent)\b|\b(terminate|termination|consent)\b.{0,180}\bchange of control\b", "Transaction consent or termination trigger", "An acquisition or financing may require counterparty consent or allow termination.", "Confirm consent requirements and obtain a transaction-specific waiver or consent.", "Corporate development and legal"),
    ("Change of Control / Assignment", "AMBER", r"\bassignment\b.{0,180}\bprior written consent\b", "Assignment requires consent", "A transfer of the agreement may be restricted after a transaction.", "Check whether the transaction structure triggers this assignment restriction.", "Legal"),
    ("Termination / Renewal", "RED", r"\bimmediate termination\b", "Immediate termination right", "The counterparty may have an unusually fast exit route.", "Validate trigger scope and revenue/transition exposure.", "Commercial and legal"),
    ("Termination / Renewal", "AMBER", r"\btermination for convenience\b", "Termination for convenience", "Revenue may be cancellable before the planned contract term.", "Model churn and committed-cost exposure; consider notice and wind-down protection.", "Commercial"),
    ("Termination / Renewal", "AMBER", r"\bauto[- ]renew", "Auto-renewal deadline", "A missed notice date can lock in or extend commercial commitments.", "Record the notice deadline and owner in the contract calendar.", "Commercial operations"),
    ("Revenue / Payment", "RED", r"\bnet\s+(?:9\d|[1-9]\d{2,})\b", "Extended payment terms", "Long payment timing can reduce cash conversion and working-capital quality.", "Quantify DSO impact and seek shorter terms where possible.", "Finance and commercial"),
    ("Revenue / Payment", "AMBER", r"\bnet\s+(?:[6-8]\d)\b", "Long payment terms", "Payment terms exceed a typical short-cycle commercial position.", "Confirm margin and cash-flow impact.", "Finance"),
    ("Pricing / MFN / Exclusivity", "RED", r"\bexclusiv", "Exclusivity restriction", "Exclusivity can constrain future customers, channels, or product strategy.", "Confirm scope, duration, carve-outs, and revenue trade-off.", "Commercial leadership and legal"),
    ("Pricing / MFN / Exclusivity", "AMBER", r"\b(most favou?red|MFN)\b", "Most-favoured-customer pricing", "MFN rights can compress margins or require future repricing.", "Identify affected products, customers, and pricing mechanics.", "Commercial and finance"),
    ("Limitation of Liability", "RED", r"\b(unlimited|uncapped)\b.{0,160}\bliabilit|\bliabilit.{0,160}\b(unlimited|uncapped)\b", "Potentially uncapped liability", "Loss exposure may exceed the expected economics of the agreement.", "Escalate cap and carve-out structure for legal and insurance review.", "Legal and insurance"),
    ("Limitation of Liability", "AMBER", r"\bliability\b.{0,160}\b(?:five|5)\s*(?:x|times)\b", "Elevated liability cap", "A high cap may create exposure disproportionate to contract value.", "Compare cap to revenue, insurance, and deal-risk tolerance.", "Legal and finance"),
    ("Indemnity", "AMBER", r"\bindemnif", "Indemnity obligation", "Indemnity can shift third-party, IP, privacy, or operational loss exposure.", "Confirm scope, defence control, exclusions, and cap interaction.", "Legal"),
    ("Intellectual Property", "RED", r"\b(customer|client)\b.{0,180}\b(background IP|pre[- ]existing|tools|methodolog)", "Background IP ownership risk", "Provider reusable technology or pre-existing IP may be exposed to ownership claims.", "Preserve provider background IP and reusable tools expressly.", "IP counsel"),
    ("Data Protection / Security", "RED", r"\b(unlimited|uncapped)\b.{0,160}\b(data breach|security incident|personal data)\b|\b(data breach|security incident|personal data)\b.{0,160}\b(unlimited|uncapped)\b", "Uncapped data incident exposure", "Privacy/security events may carry uncapped financial exposure.", "Escalate privacy, security, insurance, and liability cap review.", "Privacy, security, and legal"),
    ("Data Protection / Security", "AMBER", r"\b(?:24|twenty[- ]four)\s*hours?\b.{0,160}\b(?:breach|incident|security)\b", "24-hour incident notice", "A short notice obligation may be difficult to meet operationally.", "Validate incident response capability and notification trigger.", "Security and privacy"),
    ("Audit / Compliance", "RED", r"\bunlimited\b.{0,160}\baudit|\baudit.{0,160}\bunlimited\b", "Unlimited audit right", "Repeated or broad audits can disrupt operations and expose confidential systems.", "Limit audit frequency, scope, notice, and assessor access.", "Security and legal"),
    ("Service Levels / Credits", "RED", r"\buncapped\b.{0,160}\b(?:service credit|SLA)|\b(?:service credit|SLA)\b.{0,160}\buncapped\b", "Uncapped service credits", "Remedies may exceed a predictable service-credit exposure.", "Cap aggregate credits and align them with sole-remedy language.", "Commercial and legal"),
    ("Restrictive Covenants", "AMBER", r"\bnon[- ]solicit", "Non-solicitation restriction", "Hiring and staffing flexibility may be constrained.", "Confirm duration, covered people, and customary carve-outs.", "HR and legal"),
]

# Narrow, evidence-linked edge-case rules for rapid deal screening.
RULES.extend([
    ("Change of Control / Assignment", "RED", r"\b(?:assignment|assign(?:ment)?|transfer)\b.{0,180}\b(?:by operation of law|change of control)\b|\b(?:by operation of law|change of control)\b.{0,180}\b(?:assignment|assign(?:ment)?|transfer)\b", "Change-of-control transfer restriction", "A transaction may be treated as an assignment even without a conventional asset transfer.", "Confirm whether indirect change of control or transfer by operation of law needs consent, and obtain a targeted waiver if required.", "Corporate development and legal"),
    ("Pricing / MFN / Exclusivity", "RED", r"\b(?:unilateral|sole discretion)\b.{0,180}\b(?:price|pricing|fee|charge)\b|\b(?:price|pricing|fee|charge)\b.{0,180}\b(?:unilateral|sole discretion)\b", "Unilateral pricing control", "One party may be able to change economics without a negotiated approval mechanism.", "Confirm notice, caps, customer exit rights, and whether the commercial model tolerates unilateral price changes.", "Commercial leadership and finance"),
    ("Data Protection / Security", "AMBER", r"\b(?:data residency|data localisation|data localization|cross[- ]border transfer)\b", "Data location or transfer constraint", "Data-hosting or transfer restrictions can affect operating model, integration, and transaction diligence.", "Validate hosting locations, transfer mechanism, subcontractors, and compliance ownership.", "Privacy, security, and legal"),
    ("Revenue / Payment", "AMBER", r"\b(?:set[- ]off|withhold|withholding)\b.{0,180}\b(?:payment|invoice|fee|charge)\b|\b(?:payment|invoice|fee|charge)\b.{0,180}\b(?:set[- ]off|withhold|withholding)\b", "Payment set-off or withholding right", "Broad set-off or withholding can reduce payment certainty and distort revenue collection.", "Confirm the scope of disputed amounts, notice, cure process, and whether undisputed sums remain payable.", "Finance and commercial"),
])


def extract_txt(raw: bytes) -> str:
    text = raw.decode("utf-8", errors="replace").strip()
    if not text:
        raise ValueError("No text was found in the TXT file.")
    return text


TERM_PATTERNS = [
    ("Payment terms", r"\bnet\s+\d+\b"),
    ("Contract term", r"\b(?:initial )?term\s*(?:of|:)?\s*[^.;\n]{0,120}"),
    ("Renewal", r"\b(?:auto[- ]renew(?:al)?|renewal)\b[^.;\n]{0,160}"),
    ("Liability cap", r"\b(?:liability cap|aggregate liability|liability shall not exceed)\b[^.;\n]{0,180}"),
    ("Governing law", r"\bgoverning law\b[^.;\n]{0,160}"),
]


@dataclass
class Clause:
    reference: str
    title: str
    text: str
    page: int | None


@dataclass
class Signal:
    category: str
    severity: Literal["GREEN", "AMBER", "RED"]
    title: str
    business_impact: str
    recommended_action: str
    escalation: str
    clause_reference: str
    source_text: str
    page_number: int | None


def extract_pdf(raw: bytes) -> str:
    pages = []
    for page_no, page in enumerate(PdfReader(io.BytesIO(raw)).pages, start=1):
        text = page.extract_text() or ""
        if text.strip():
            pages.append(f"[Page {page_no}]\n{text}")
    if not pages:
        raise ValueError("No selectable text was found. OCR a scanned PDF before running the radar.")
    return "\n\n".join(pages)


def extract_docx(raw: bytes) -> str:
    text = "\n\n".join(p.text for p in Document(io.BytesIO(raw)).paragraphs if p.text.strip())
    if not text:
        raise ValueError("No text was found in the DOCX file.")
    return text


def page_at(text: str, position: int) -> int | None:
    markers = list(re.finditer(r"\[Page (\d+)\]", text[: position + 1]))
    return int(markers[-1].group(1)) if markers else 1


def quote(text: str, start: int, end: int, limit: int = 550) -> str:
    left = max(0, text.rfind("\n", max(0, start - limit // 2), start))
    right = text.find("\n", end, min(len(text), end + limit // 2))
    right = len(text) if right < 0 else right
    snippet = text[left:right].strip()
    return snippet[:limit]


def split_clauses(text: str) -> list[Clause]:
    """Recognise common legal numbering and avoid treating an irregular agreement as one clause."""
    heading = re.compile(
        r"(?mi)^\s*(?:(?:section|clause|article)\s+)?(?P<number>\d+(?:\.\d+){0,5}|[IVXLC]+|[A-Z])(?:\s*[.:)\-])?\s+(?P<title>[A-Z][^\n]{2,120})\s*$"
    )
    matches = list(heading.finditer(text))
    if not matches:
        paragraphs = [part.strip() for part in re.split(r"\n\s*\n", text) if len(part.strip()) >= 80]
        if len(paragraphs) <= 1:
            return [Clause("Unnumbered", "Agreement text", text, page_at(text, 0))]
        clauses, cursor = [], 0
        for index, paragraph in enumerate(paragraphs, start=1):
            position = text.find(paragraph, cursor)
            cursor = position + len(paragraph)
            clauses.append(Clause(f"Unnumbered {index}", paragraph[:70].split(".")[0], paragraph, page_at(text, position)))
        return clauses

    clauses = []
    if text[:matches[0].start()].strip():
        clauses.append(Clause("Preamble", "Agreement preamble", text[:matches[0].start()].strip(), page_at(text, 0)))
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        body = text[match.start():end].strip()
        if len(body) >= 30:
            clauses.append(Clause(match.group("number").rstrip("."), match.group("title").strip(), body, page_at(text, match.start())))
    return clauses

def categorize(clause: Clause) -> str:
    for category, pattern in CATEGORY_PATTERNS.items():
        if re.search(pattern, f"{clause.title}\n{clause.text}", re.IGNORECASE):
            return category
    return "Miscellaneous"


def make_signal(rule: tuple, clause: Clause, match: re.Match) -> Signal:
    category, severity, pattern, title, impact, action, escalation = rule
    return Signal(category, severity, title, impact, action, escalation, clause.reference, quote(clause.text, match.start(), match.end()), clause.page)


def analyze(text: str) -> tuple[list[Clause], list[Signal], list[dict]]:
    clauses = split_clauses(text)
    signals: list[Signal] = []
    terms: list[dict] = []

    for clause in clauses:
        for rule in RULES:
            match = re.search(rule[2], clause.text, re.IGNORECASE | re.DOTALL)
            if match:
                signals.append(make_signal(rule, clause, match))
        for label, pattern in TERM_PATTERNS:
            match = re.search(pattern, clause.text, re.IGNORECASE)
            if match:
                terms.append({
                    "label": label,
                    "value": quote(clause.text, match.start(), match.end(), 180),
                    "clause_reference": clause.reference,
                    "source_text": quote(clause.text, match.start(), match.end()),
                    "page_number": clause.page,
                    "category": categorize(clause),
                })

    unique_signals = list({(s.category, s.title, s.source_text): s for s in signals}.values())
    unique_terms = list({(t["label"], t["source_text"]): t for t in terms}.values())
    return clauses, unique_signals, unique_terms


def score(signals: list[Signal]) -> dict:
    weights = {"RED": 20, "AMBER": 7, "GREEN": 0}
    counts = {level: sum(s.severity == level for s in signals) for level in weights}
    value = max(0, 100 - sum(weights[s.severity] for s in signals))
    rating = "RED" if counts["RED"] >= 2 or value < 55 else "AMBER" if counts["RED"] or counts["AMBER"] >= 3 or value < 80 else "GREEN"
    order = {"RED": 0, "AMBER": 1, "GREEN": 2}
    priorities = sorted(signals, key=lambda s: (order[s.severity], s.category))
    return {"rating": rating, "score": value, "counts": counts, "priority_signals": priorities}


def memo(mode: str, scorecard: dict) -> dict:
    signals = scorecard["priority_signals"]
    rating = scorecard["rating"]
    decision = "PAUSE FOR DILIGENCE" if rating == "RED" else "PROCEED WITH CONDITIONS" if rating == "AMBER" else "PROCEED"
    framing = "investment case" if mode == "Investment diligence" else "commercial approval"
    priorities = [f"{s.title}: {s.recommended_action}" for s in signals[:5]]
    questions = [f"Can the team confirm and evidence the mitigation for: {s.title}?" for s in signals if s.severity in {"RED", "AMBER"}][:6]
    actions = list(dict.fromkeys(s.recommended_action for s in signals if s.severity in {"RED", "AMBER"}))[:6]
    return {
        "decision": decision,
        "headline": f"{rating} deal-risk rating based on {len(signals)} evidence-backed signal(s).",
        "investment_or_commercial_impact": f"The {framing} should account for the highlighted revenue, transferability, liability, margin, and operational exposures.",
        "top_priorities": priorities or ["No rule-triggered deal signals were found; counsel should still review material provisions."],
        "diligence_questions": questions or ["Confirm whether any material commercial terms exist outside the uploaded agreement."],
        "commercial_actions": actions or ["Retain evidence and complete qualified legal review."],
        "lawyer_review_note": "Rule matches are screening prompts, not conclusions. Verify exact wording, applicability, jurisdiction, and commercial context.",
    }


def badge(level: str) -> str:
    color = {"RED": "#b42318", "AMBER": "#b54708", "GREEN": "#027a48"}.get(level, "#344054")
    return f"<span style='color:{color};font-weight:700'>{level}</span>"


def report_markdown(result: dict) -> str:
    scorecard, decision = result["scorecard"], result["memo"]
    lines = [f"# {APP_NAME}", "", f"> {DISCLAIMER}", "", "## Decision", f"**{decision['decision']}**", decision["headline"], "", "## Signals"]
    for signal in result["signals"]:
        lines += [f"### {signal['severity']} — {signal['title']}", signal["business_impact"], f"Action: {signal['recommended_action']}", f"Evidence: {signal['source_text']}", ""]
    return "\n".join(lines)


def main() -> None:
    st.set_page_config(page_title=APP_NAME, page_icon="📡", layout="wide")
    st.title(APP_NAME)
    st.caption("Fast, local, deterministic screening for PE/VC, corporate development, and commercial approvals")
    st.warning(DISCLAIMER, icon="⚠️")

    with st.sidebar:
        st.header("Review setup")
        mode = st.radio("Decision context", ["Investment diligence", "Commercial approval"])
        upload = st.file_uploader("Contract or agreement", type=["pdf", "docx", "txt"])
        run = st.button("Run local deal radar", type="primary", disabled=upload is None)
        st.caption("Runs fully in Python: no API key, network call, Ollama, or LLM.")

    if "result" not in st.session_state:
        st.session_state.result = None
    if "human_decisions" not in st.session_state:
        st.session_state.human_decisions = {}

    if upload and st.session_state.get("active_upload") != upload.name:
        st.session_state.result = None
        st.session_state.human_decisions = {}
        st.session_state.active_upload = upload.name

    if run and upload:
        try:
            raw = upload.getvalue()
            text = extract_pdf(raw) if upload.name.lower().endswith(".pdf") else extract_docx(raw) if upload.name.lower().endswith(".docx") else extract_txt(raw)
            with st.spinner("Parsing clauses and applying local commercial-risk rules…"):
                clauses, signals, terms = analyze(text)
                scorecard = score(signals)
                st.session_state.result = {
                    "clauses": [asdict(c) for c in clauses],
                    "signals": [asdict(s) for s in signals],
                    "terms": terms,
                    "scorecard": {**scorecard, "priority_signals": [asdict(s) for s in scorecard["priority_signals"]]},
                    "memo": memo(mode, scorecard),
                    "coverage": len(text),
                }
            st.success("Local deal radar complete.")
        except Exception as exc:
            st.session_state.result = None
            st.error(str(exc))

    result = st.session_state.result
    if not result:
        st.info("Upload a text-based PDF, DOCX, or TXT agreement. The radar will run instantly on local Python rules.")
        return

    scorecard, decision = result["scorecard"], result["memo"]
    st.caption(f"Full-document coverage: {result['coverage']:,} characters parsed locally; no text was truncated.")
    memo_tab, signals_tab, terms_tab, clauses_tab, human_tab = st.tabs(["Decision memo", "Deal signals", "Key terms", "Clauses", "Human review"])

    with memo_tab:
        left, right = st.columns([1, 2])
        with left:
            st.metric("Deal risk rating", scorecard["rating"])
            st.metric("Radar score", f"{scorecard['score']}/100")
            counts = scorecard["counts"]
            st.write(f"🔴 {counts['RED']} red · 🟠 {counts['AMBER']} amber · 🟢 {counts['GREEN']} green")
        with right:
            st.subheader(decision["decision"])
            st.write(decision["headline"])
            st.write(decision["investment_or_commercial_impact"])
        for label, items in [("Top priorities", decision["top_priorities"]), ("Diligence questions", decision["diligence_questions"]), ("Commercial actions", decision["commercial_actions"])]:
            st.subheader(label)
            for item in items:
                st.write(f"- {item}")
        st.info(decision["lawyer_review_note"])
        st.download_button("Download radar report", report_markdown(result), "deal_contract_radar.md", "text/markdown")

    with signals_tab:
        categories = sorted({s["category"] for s in result["signals"]})
        category_filter = st.multiselect("Filter categories", categories)
        severity_filter = st.multiselect("Filter severity", ["RED", "AMBER", "GREEN"], default=["RED", "AMBER", "GREEN"])
        for signal in result["signals"]:
            if category_filter and signal["category"] not in category_filter:
                continue
            if signal["severity"] not in severity_filter:
                continue
            st.markdown(f"{badge(signal['severity'])} — **{signal['title']}**", unsafe_allow_html=True)
            st.write(signal["business_impact"])
            st.write(f"**Recommended action:** {signal['recommended_action']}")
            st.caption(f"Escalation: {signal['escalation']} · Clause {signal['clause_reference']} · Page {signal['page_number'] or 'not available'}")
            with st.expander("Contract evidence"):
                st.code(signal["source_text"], language=None)
            st.divider()

    with terms_tab:
        for term in result["terms"]:
            with st.expander(f"{term['label']}: {term['value']}"):
                st.caption(f"Clause {term['clause_reference']} · Page {term['page_number'] or 'not available'}")
                st.code(term["source_text"], language=None)

    with clauses_tab:
        for clause in result["clauses"]:
            with st.expander(f"{clause['reference']} — {clause['title']}"):
                st.caption(f"Page {clause['page'] or 'not available'} · {categorize(Clause(**clause))}")
                st.code(clause["text"], language=None)

    with human_tab:
        st.caption("The app cannot accept, reject, or modify a contract. Decisions below are human-only.")
        for index, signal in enumerate(result["signals"]):
            key = f"{signal['category']}:{signal['title']}:{index}"
            st.markdown(f"**{signal['title']}**")
            decision_value = st.selectbox("Human decision", ["Pending", "Proceed", "Diligence further", "Negotiate", "Escalate", "Ignore"], key=f"decision_{index}")
            note = st.text_input("Decision note", key=f"note_{index}")
            st.session_state.human_decisions[key] = {"decision": decision_value, "note": note}
            st.divider()


if __name__ == "__main__":
    main()
