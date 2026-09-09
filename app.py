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
    ("Pricing / MFN / Exclusivity", "RED", r"\b(?:shall|will|must|agrees? to|is granted|appoint(?:ed)? as|grant(?:s|ed)?)\b.{0,180}\b(?:exclusive|exclusivity)\b|\b(?:exclusive|exclusivity)\b.{0,180}\b(?:shall|will|must|agrees? to|is granted|appoint(?:ed)? as|grant(?:s|ed)?)\b", "Exclusivity restriction", "Exclusivity can constrain future customers, channels, or product strategy.", "Confirm scope, duration, carve-outs, and revenue trade-off.", "Commercial leadership and legal"),
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


RECITAL_TITLES = {"whereas", "recitals", "preamble", "background"}


def is_recital_clause(clause: Clause) -> bool:
    title = clause.title.strip().lower().rstrip(":")
    opening = clause.text.lstrip().lower()
    return title in RECITAL_TITLES or opening.startswith("whereas") or opening.startswith("recitals")


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
    clause_title: str
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


STRUCTURAL_TITLE_WORDS = {
    "definitions", "scope", "services", "fees", "payment", "pricing", "term", "termination",
    "renewal", "liability", "indemnity", "confidentiality", "data protection", "security",
    "intellectual property", "audit", "compliance", "assignment", "subcontracting",
    "governing law", "dispute resolution", "notices", "insurance", "service levels",
    "restrictions", "publicity", "warranties", "general", "miscellaneous",
}


def heading_candidates(text: str) -> list[tuple[int, str, str]]:
    """Find numbered, split-line, and visual legal headings from plain extracted text."""
    numbered = re.compile(
        r"(?mi)^\s*(?:(?:section|clause|article)\s+)?(?P<number>\d+(?:\.\d+){0,5}|[IVXLC]+|[A-Z])(?:\s*[.:)\-])?\s+(?P<title>[^\n]{3,120})\s*$"
    )
    candidates: dict[int, tuple[int, str, str]] = {}
    for match in numbered.finditer(text):
        title = match.group("title").strip()
        if len(title.split()) <= 16 and not re.search(r"[.;!?]$", title):
            candidates[match.start()] = (match.start(), match.group("number").rstrip("."), title)

    lines = list(re.finditer(r"(?m)^([^\n]+)$", text))
    number_only = re.compile(r"(?i)^\s*(?:(?:section|clause|article)\s+)?(?P<number>\d+(?:\.\d+){0,5}|[IVXLC]+|[A-Z])(?:\s*[.:)\-])?\s*$")
    for index, line in enumerate(lines):
        position, value = line.start(), line.group(1).strip()
        split_number = number_only.match(value)
        if split_number:
            for later in lines[index + 1:]:
                title = later.group(1).strip()
                if not title:
                    continue
                if 3 <= len(title) <= 120 and len(title.split()) <= 16 and not re.search(r"[.;!?]$", title):
                    candidates[position] = (position, split_number.group("number").rstrip("."), title)
                break
            continue
        if position in candidates or not (3 <= len(value) <= 100) or len(value.split()) > 12:
            continue
        if re.search(r"[.;!?]$", value) or value.startswith("[Page "):
            continue
        lower = value.lower()
        words = [word for word in re.findall(r"[A-Za-z]+", value) if word]
        title_case_ratio = sum(word[0].isupper() for word in words) / len(words) if words else 0
        known_heading = any(term in lower for term in STRUCTURAL_TITLE_WORDS)
        letters = re.sub(r"[^A-Za-z]", "", value)
        all_caps = bool(letters) and letters.isupper()
        next_text = ""
        for later in lines[index + 1:]:
            candidate = later.group(1).strip()
            if candidate:
                next_text = candidate
                break
        if (known_heading or all_caps or title_case_ratio >= 0.80) and len(next_text) >= 35:
            candidates[position] = (position, "Heading-derived", value)
    return [candidates[key] for key in sorted(candidates)]


def derived_fragments(text: str) -> list[Clause]:
    """Create traceable fragments only where the source supplies no usable provision boundaries."""
    fragments, cursor, counter = [], 0, 0
    paragraphs = [part.strip() for part in re.split(r"\n\s*\n", text) if len(part.strip()) >= 80]
    for paragraph in paragraphs:
        position = text.find(paragraph, cursor)
        cursor = position + len(paragraph)
        sentences = re.split(r"(?<=[.;!?])\s+(?=[A-Z])", paragraph)
        buffer = ""
        for sentence in sentences:
            if buffer and len(buffer) + len(sentence) + 1 > 850:
                counter += 1
                fragments.append(Clause(f"Derived fragment {counter}", buffer[:85].split(".")[0], buffer, page_at(text, position)))
                position += len(buffer)
                buffer = sentence
            else:
                buffer = (buffer + " " + sentence).strip()
        if len(buffer) >= 80:
            counter += 1
            fragments.append(Clause(f"Derived fragment {counter}", buffer[:85].split(".")[0], buffer, page_at(text, position)))
    return fragments or [Clause("Derived fragment 1", "Agreement text", text, page_at(text, 0))]


def split_clauses(text: str) -> list[Clause]:
    headings = heading_candidates(text)
    if not headings:
        return derived_fragments(text)
    clauses = []
    first_start = headings[0][0]
    if text[:first_start].strip():
        clauses.append(Clause("Preamble", "Agreement preamble", text[:first_start].strip(), page_at(text, 0)))
    heading_counts = {"Heading-derived": 0}
    for index, (start, reference, title) in enumerate(headings):
        end = headings[index + 1][0] if index + 1 < len(headings) else len(text)
        body = text[start:end].strip()
        if len(body) < 30:
            continue
        if reference == "Heading-derived":
            heading_counts["Heading-derived"] += 1
            reference = f"Heading-derived {heading_counts['Heading-derived']}"
        clauses.append(Clause(reference, title, body, page_at(text, start)))
    return clauses or derived_fragments(text)


def structure_summary(clauses: list[Clause]) -> dict[str, int]:
    return {
        "numbered": sum(bool(re.match(r"^(?:\d|[IVXLC]+$|[A-Z]$)", clause.reference)) for clause in clauses),
        "heading_derived": sum(clause.reference.startswith("Heading-derived") for clause in clauses),
        "derived": sum(clause.reference.startswith("Derived fragment") for clause in clauses),
    }

def categorize(clause: Clause) -> str:
    for category, pattern in CATEGORY_PATTERNS.items():
        if re.search(pattern, f"{clause.title}\n{clause.text}", re.IGNORECASE):
            return category
    return "Miscellaneous"


def make_signal(rule: tuple, clause: Clause, match: re.Match) -> Signal:
    category, severity, pattern, title, impact, action, escalation = rule
    return Signal(category, severity, title, impact, action, escalation, clause.reference, clause.title, quote(clause.text, match.start(), match.end()), clause.page)


def analyze(text: str) -> tuple[list[Clause], list[Signal], list[dict]]:
    clauses = split_clauses(text)
    signals: list[Signal] = []
    terms: list[dict] = []

    for clause in clauses:
        if not is_recital_clause(clause):
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
                    "clause_title": clause.title,
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


MEMO_GUIDANCE = {
    "Change of Control / Assignment": ("Does the contemplated transaction, financing, or internal reorganisation trigger a consent, notice, or termination right?", "Corporate development and legal: map the transaction structure against the exact assignment/change-of-control wording and seek a targeted waiver if needed."),
    "Termination / Renewal": ("What event, notice period, cure right, or renewal date could interrupt the forecast revenue or service relationship?", "Commercial owner and legal: calendar notice dates, quantify wind-down exposure, and negotiate the identified exit mechanism."),
    "Revenue / Payment": ("What is the cash-flow effect of the payment term, dispute mechanism, set-off, or withholding right?", "Finance and commercial: model DSO and collection exposure, then set a clear due-date and disputed-sums process."),
    "Pricing / MFN / Exclusivity": ("What products, customers, channels, or future prices are constrained, and for how long?", "Commercial leadership: quantify the revenue trade-off and narrow the pricing, MFN, or exclusivity perimeter."),
    "Limitation of Liability": ("Does the cap and its carve-outs align with the deal economics, insurance, and plausible loss scenarios?", "Legal and finance: set a negotiated aggregate cap, loss exclusions, and scoped carve-outs."),
    "Indemnity": ("Which third-party claims are covered, who controls defence and settlement, and how does the obligation interact with the cap?", "Legal: limit covered claims and document defence, settlement, exclusion, and cap mechanics."),
    "Intellectual Property": ("Are ownership, licence, and background-IP rights sufficient for the transaction without exposing reusable technology?", "IP counsel: separate background IP, deliverables, embedded materials, and permitted use rights."),
    "Data Protection / Security": ("Can the operating model meet the stated data, incident, hosting, and security commitment?", "Privacy and security: validate roles, incident timeline, transfer requirements, controls, and remediation ownership."),
    "Audit / Compliance": ("Are audit scope, frequency, cost, system access, and remediation obligations proportionate?", "Security and legal: replace unrestricted access with defined notice, frequency, scope, and assurance-report mechanics."),
    "Service Levels / Credits": ("Are metrics, exclusions, credits, and chronic-failure remedies commercially sustainable?", "Commercial and legal: define measurement, exclusions, aggregate credit cap, and chronic-failure path."),
    "Restrictive Covenants": ("Who is restricted, for what period, and are customary staffing or publicity carve-outs preserved?", "HR and legal: narrow the covered population, duration, and prohibited conduct and add standard carve-outs."),
    "Dispute Resolution / Governing Law": ("Will the forum and escalation process support timely, cost-effective enforcement for this deal?", "Legal: confirm governing law, forum, escalation, costs, and urgent-relief rights."),
}


def distinct_priority_signals(signals: list[Signal], limit: int = 5) -> list[Signal]:
    selected, seen_categories = [], set()
    for signal in signals:
        if signal.category not in seen_categories:
            selected.append(signal)
            seen_categories.add(signal.category)
        if len(selected) >= limit:
            break
    return selected


def memo(mode: str, scorecard: dict) -> dict:
    signals = scorecard["priority_signals"]
    rating = scorecard["rating"]
    decision = "PAUSE FOR DILIGENCE" if rating == "RED" else "PROCEED WITH CONDITIONS" if rating == "AMBER" else "PROCEED"
    framing = "investment case" if mode == "Investment diligence" else "commercial approval"
    priorities = distinct_priority_signals(signals)
    top_priorities = [f"{signal.severity} · {signal.category} · Clause {signal.clause_reference} — {signal.clause_title}: {signal.business_impact}" for signal in priorities]
    questions = []
    actions = []
    for signal in priorities:
        question, action = MEMO_GUIDANCE.get(signal.category, ("What operative obligation, exception, remedy, and commercial impact does this evidence create?", "Legal and commercial: verify the source wording and document a negotiated mitigation."))
        if question not in questions:
            questions.append(question)
        if action not in actions:
            actions.append(action)
    return {
        "decision": decision,
        "headline": f"{rating} deal-risk rating based on {len(signals)} evidence-backed signal(s) across {len({signal.category for signal in signals})} risk category/categories.",
        "investment_or_commercial_impact": f"The {framing} should be assessed against the distinct highest-priority categories below; repeated signals within the same category remain included in the score.",
        "top_priorities": top_priorities or ["No rule-triggered deal signals were found; counsel should still review material provisions."],
        "diligence_questions": questions or ["Confirm whether any material commercial terms exist outside the uploaded agreement."],
        "commercial_actions": actions or ["Retain evidence and complete qualified legal review."],
        "lawyer_review_note": "Each priority identifies the matching clause reference and title. Rule matches are screening prompts, not conclusions; verify wording, applicability, jurisdiction, and commercial context.",
    }


def badge(level: str) -> str:
    color = {"RED": "#b42318", "AMBER": "#b54708", "GREEN": "#027a48"}.get(level, "#344054")
    return f"<span style='color:{color};font-weight:700'>{level}</span>"


DRAFTING_PLAYBOOK = {
    "Change of Control / Assignment": (
        "Preserve transaction flexibility while allowing consent only for a materially adverse replacement counterparty.",
        "[Counterparty] shall not unreasonably withhold, condition, or delay consent to an assignment by [Company] to an Affiliate or in connection with a merger, reorganisation, sale of substantially all assets, or change of control, provided that the assignee assumes this Agreement in writing. No consent is required for an internal reorganisation that does not reduce [Counterparty]'s contractual protections.",
        "Define whether indirect change of control is covered, identify any regulated-consent exception, and agree a response deadline and deemed-consent consequence where appropriate.",
    ),
    "Termination / Renewal": (
        "Make exit rights, cure opportunities, renewal notice, and commercial consequences predictable.",
        "Neither Party may terminate this Agreement for material breach unless it first gives written notice describing the breach in reasonable detail and the breaching Party fails to cure it within [30] days; provided that a non-curable breach may be terminated on written notice. Any termination for convenience requires at least [90] days' prior written notice and does not relieve [Counterparty] of accrued payment obligations or agreed wind-down charges.",
        "Tailor cure periods, convenience rights, committed spend, transition support, data return, and survival obligations to the transaction.",
    ),
    "Revenue / Payment": (
        "Protect collection of undisputed amounts and prevent broad withholding from becoming a cash-flow risk.",
        "Amounts properly invoiced and not disputed in good faith are due within [30] days of receipt. A Party may withhold only the specific portion of an invoice that it disputes in writing before the due date, with reasonable detail; all undisputed amounts remain payable when due. The Parties will work in good faith to resolve a disputed amount within [15] days.",
        "Set the currency, tax treatment, late charge, suspension threshold, invoice requirements, and any permitted set-off rights.",
    ),
    "Pricing / MFN / Exclusivity": (
        "Prevent open-ended pricing, MFN, or exclusivity commitments from constraining future economics.",
        "Fees may be changed only by a written amendment signed by both Parties, except for an annual increase not exceeding [X%] on at least [60] days' prior notice. Any most-favoured-customer or exclusivity commitment applies only to the expressly identified products, territory, customer segment, and period, and excludes promotional, bundled, legacy, and materially different transactions.",
        "Define pricing comparators, the duration and carve-outs of any restriction, termination rights, and the commercial consideration for exclusivity.",
    ),
    "Limitation of Liability": (
        "Set a measurable aggregate exposure ceiling and expressly negotiated carve-outs.",
        "Except for the Excluded Claims, each Party's aggregate liability arising out of or relating to this Agreement will not exceed the fees paid or payable under this Agreement in the [12] months preceding the event giving rise to liability. Neither Party will be liable for indirect, incidental, special, consequential, exemplary, or punitive damages, or lost profits, revenue, goodwill, or data, except to the extent such amounts are payable to a third party under an agreed indemnity.",
        "Define the cap base, claim period, excluded-loss treatment, insurance alignment, and any carve-outs for confidentiality, IP, data, fraud, or wilful misconduct.",
    ),
    "Indemnity": (
        "Limit indemnity to a defined set of third-party claims and establish defence and settlement controls.",
        "[Indemnifying Party] will defend and indemnify [Indemnified Party] against final third-party claims to the extent arising from [defined IP infringement / bodily injury / property damage], provided that [Indemnified Party] promptly notifies [Indemnifying Party], reasonably cooperates, and permits [Indemnifying Party] to control the defence. No settlement may impose liability, admission, or non-monetary obligation on [Indemnified Party] without its prior written consent.",
        "Specify exclusions, IP remediation options, control of counsel, notice prejudice, cap interaction, and whether data/privacy claims are included.",
    ),
    "Intellectual Property": (
        "Protect background technology and clearly allocate deliverable rights.",
        "Each Party retains all right, title, and interest in its pre-existing materials, tools, software, methodologies, and know-how. To the extent [Company] incorporates its background materials into a deliverable, [Counterparty] receives a non-exclusive, worldwide, perpetual licence to use those materials solely as embedded in and necessary to use the deliverable. No implied licence is granted.",
        "Separate background IP, deliverables, customer materials, open-source components, feedback, and residual know-how.",
    ),
    "Data Protection / Security": (
        "Convert data and security obligations into operationally achievable, measurable commitments.",
        "[Processor/Service Provider] will process Personal Data only on documented instructions, implement appropriate technical and organisational measures, notify [Controller/Customer] without undue delay and in any event within [X] hours after confirming a Security Incident, and return or delete Personal Data at termination unless retention is legally required. Subprocessors require prior [notice/consent] and written obligations no less protective than this Agreement.",
        "Tailor roles, incident trigger and timeline, hosting location, transfer mechanism, audit evidence, regulatory assistance, and allocation of remediation costs.",
    ),
    "Audit / Compliance": (
        "Make audit rights proportionate, secure, and non-disruptive.",
        "[Customer] may, no more than once in any [12]-month period and on at least [30] days' prior written notice, audit [Company]'s compliance with the applicable obligations during normal business hours, subject to confidentiality, reasonable security requirements, and no access to other customers' information. Independent assurance reports may satisfy the audit requirement unless a material non-compliance is reasonably suspected.",
        "Set frequency, scope, assessor qualifications, costs, notice, remediation, and treatment of sensitive systems and data.",
    ),
    "Service Levels / Credits": (
        "Tie performance remedies to defined metrics and cap financial exposure.",
        "Service credits are [Customer]'s sole and exclusive monetary remedy for a Service Level failure. Aggregate service credits in any calendar month will not exceed [X%] of the monthly fees for the affected Service. Service Levels exclude failures caused by [Customer], third-party systems, scheduled maintenance, force majeure, or events outside [Company]'s reasonable control.",
        "Define metrics, measurement method, exclusions, credit calculation, escalation, chronic-failure rights, and interaction with termination remedies.",
    ),
    "Restrictive Covenants": (
        "Narrow staffing, publicity, or competitive restrictions to what is commercially necessary.",
        "During the Term and for [12] months thereafter, neither Party will knowingly solicit for employment an employee of the other Party who was materially involved in performing this Agreement, except through general solicitations not targeted at that employee, responses to unsolicited applications, or use of recruiters not directed to target that employee.",
        "Define covered people, duration, geography, general-solicitation and recruiter carve-outs, and any agreed remedy.",
    ),
    "Dispute Resolution / Governing Law": (
        "Create a predictable escalation route while preserving urgent-relief rights.",
        "The Parties will first escalate any dispute to designated business representatives for [15] days before commencing formal proceedings. This Agreement is governed by the laws of [jurisdiction], and the courts of [forum] have exclusive jurisdiction; either Party may seek interim injunctive relief in any court of competent jurisdiction.",
        "Choose governing law, forum or arbitration rules, language, cost allocation, notice mechanics, and exceptions for urgent relief.",
    ),
}


SIGNAL_DRAFTING_OVERRIDES = {
    "Exclusivity restriction": (
        "Convert any exclusivity concept into a narrowly scoped, expressly bargained-for commercial commitment rather than a broad implied restraint.",
        "No exclusivity is granted under this Agreement except to the extent expressly stated in this clause. If exclusivity is agreed, it applies only to [identified products/services] in [territory/customer segment] during [period], is conditional on [minimum commitment/performance threshold], and does not restrict [Company]'s existing customers, Affiliates, channels, products, or opportunities outside that defined scope. Any breach remedy is limited to [specified remedy].",
        "Confirm whether exclusivity is intended at all. If it is, define beneficiary, product/service, territory, channel, duration, performance conditions, carve-outs, and the sole remedy; do not rely on descriptive recital language.",
    ),
}

SIGNAL_PARTY_OVERRIDES = {
    ("Client / Customer", "Exclusivity restriction"): (
        "If the Client is seeking exclusivity, ensure it is an enforceable commercial benefit with defined scope and a remedy for breach.",
        "Company grants Client an exclusive right to [market/distribute/use] the [identified products/services] solely within [territory/customer segment] during [period], subject to Client meeting [objective commitments]. Company will not appoint another [provider/distributor] for that defined scope. This exclusivity does not apply to [express carve-outs], and Client may terminate the exclusivity arrangement if Company materially breaches it and fails to cure within [30] days.",
    ),
    ("Company / Service Provider", "Exclusivity restriction"): (
        "Avoid an implied or unlimited restraint; if exclusivity is commercially necessary, make it conditional, narrow, and time-limited.",
        "Except for the expressly defined exclusivity in this clause, Company remains free to market, sell, license, appoint, and provide the products/services to any person. Any exclusivity is limited to [scope] in [territory] for [period], ends automatically if Client fails to meet [minimum commitment/performance threshold], and excludes Company's existing customers, Affiliates, product lines, and opportunities outside the stated scope.",
    ),
}


def drafting_response(signal: dict) -> dict[str, str]:
    objective, clause, tailoring = SIGNAL_DRAFTING_OVERRIDES.get(
        signal["title"], DRAFTING_PLAYBOOK.get(
        signal["category"],
        (
            "Make the obligation, exception, remedy, and cost/risk allocation explicit.",
            "The Parties will document the scope of the relevant obligation, the applicable exceptions, notice and cure process, remedy, and any agreed financial limitation in a written amendment signed by both Parties.",
            "Tie the wording to the exact source evidence, parties, jurisdiction, and commercial position before use.",
        ),
    ))
    return {"objective": objective, "clause": clause, "tailoring": tailoring}


def drafting_pack(result: dict) -> str:
    lines = ["# Deal Risk Radar — Drafting Responses", "", "> Sample negotiating language only; qualified counsel must tailor it to the agreement, parties, jurisdiction, and commercial position.", ""]
    for signal in result["signals"]:
        if signal["severity"] not in {"RED", "AMBER"}:
            continue
        response = drafting_response(signal)
        lines += [f"## {signal['severity']} — {signal['title']}", f"Clause: {signal['clause_reference']} — {signal['clause_title']}\nSource evidence: {signal['source_text']}", "", "### Drafting objective", response["objective"], "", "### Sample clause", response["clause"], "", "### Tailoring points", response["tailoring"], ""]
    return "\n".join(lines)


PARTY_LENS = {
    "Client / Customer": {
        "Change of Control / Assignment": ("Protect continuity and prevent a transfer to an unsuitable provider without an effective consent right.", "Client consent is required only for an assignment that materially and adversely affects the Services, and Client will not unreasonably withhold or delay consent. A change of control of Company will be treated as an assignment only if expressly stated and only where it materially impairs performance."),
        "Termination / Renewal": ("Preserve a workable exit route, transition support, and protection from unwanted renewals.", "Client may terminate for Company material breach not cured within [30] days, repeated material Service Level failures, or for convenience on [90] days' notice. On termination, Company will provide reasonable transition assistance and return Client Data in the agreed format."),
        "Revenue / Payment": ("Retain a good-faith invoice-dispute process without withholding undisputed sums.", "Client may withhold only the genuinely disputed portion of an invoice after giving written notice with reasonable detail; all undisputed amounts remain payable when due. No late charge applies to the disputed amount while the Parties work in good faith to resolve it."),
        "Pricing / MFN / Exclusivity": ("Avoid uncontrolled price changes and ensure any preferential pricing or exclusivity is specific and enforceable.", "Any fee increase requires at least [60] days' notice and will not exceed [X%] in a contract year. Any MFN or exclusivity benefit for Client applies to the specified products, territory, and term and is enforceable through a price adjustment or termination right."),
        "Limitation of Liability": ("Keep meaningful remedies for the Client's most serious loss scenarios.", "The liability cap will not apply, or will apply at an enhanced cap of [X], to Company's breach of confidentiality, data-protection obligations, IP indemnity, fraud, or wilful misconduct. The agreed cap does not prevent recovery of direct, documented remediation costs."),
        "Indemnity": ("Secure defence and reimbursement for defined third-party claims.", "Company will defend and indemnify Client against third-party claims arising from Company’s IP infringement, breach of applicable data-protection law, or gross negligence, subject to Client’s prompt notice and reasonable cooperation."),
        "Intellectual Property": ("Ensure the Client receives usable rights in deliverables and continuity if the relationship ends.", "Client owns the specifically identified Deliverables on payment, excluding Company Background IP. Company grants Client a perpetual, irrevocable, worldwide licence to use embedded Background IP as necessary to receive the benefit of the Deliverables."),
        "Data Protection / Security": ("Require actionable security, incident, and downstream-processing protections.", "Company will notify Client of a confirmed Security Incident without undue delay and no later than [X] hours, provide regular updates, cooperate with investigation and remediation, and remain responsible for its subprocessors’ compliance."),
        "Audit / Compliance": ("Obtain proportionate assurance and remediation visibility.", "Client may receive current independent assurance reports and, where reasonably necessary after a material incident or credible non-compliance concern, conduct a proportionate audit subject to confidentiality and security controls."),
        "Service Levels / Credits": ("Make performance commitments measurable and give an effective chronic-failure remedy.", "If Company misses a material Service Level in [X] months in any [Y]-month period, Client may terminate the affected Services without early-termination charge, in addition to accrued service credits."),
        "Restrictive Covenants": ("Narrow restrictions that could unnecessarily limit Client staffing or business operations.", "Any non-solicitation restriction is limited to personnel materially involved in the Services, lasts no more than [12] months, and excludes general solicitations, unsolicited applications, and independent recruiters not directed to target protected personnel."),
        "Dispute Resolution / Governing Law": ("Preserve a practical enforcement forum and urgent-relief rights.", "Nothing prevents Client from seeking interim or injunctive relief for confidentiality, data, IP, or service-continuity harm in a court of competent jurisdiction while the Parties follow the agreed escalation process."),
    },
    "Company / Service Provider": {
        "Change of Control / Assignment": ("Maintain transaction and internal-reorganisation flexibility.", "Company may assign this Agreement to an Affiliate or in connection with a merger, reorganisation, financing, or sale of all or substantially all of its assets, provided the assignee assumes Company’s obligations. Client consent may not be unreasonably withheld, conditioned, or delayed."),
        "Termination / Renewal": ("Prevent abrupt, uncompensated exits and preserve cure opportunities.", "Company receives written notice and at least [30] days to cure a material breach before termination, except for non-curable breaches. A convenience termination by Client requires [90] days' notice and payment of all undisputed accrued fees and agreed wind-down charges."),
        "Revenue / Payment": ("Protect collection certainty and restrict withholding to genuine disputes.", "Client will pay all undisputed amounts within [30] days. Any disputed amount must be identified in writing before the due date with reasonable detail, and Client may not set off or withhold any undisputed amount."),
        "Pricing / MFN / Exclusivity": ("Keep pricing discretion bounded but commercially workable and avoid overbroad MFN/exclusivity obligations.", "Company may adjust fees once per contract year on at least [60] days' notice by no more than [X%]. Any MFN or exclusivity commitment excludes bespoke, bundled, promotional, legacy, or materially different transactions and expires after [X] months."),
        "Limitation of Liability": ("Cap aggregate exposure and exclude loss categories that are not proportionate to the deal economics.", "Company's aggregate liability under this Agreement will not exceed the fees paid or payable in the [12] months preceding the event giving rise to the claim. Company will not be liable for indirect, consequential, special, punitive, or lost-profit damages, subject only to expressly negotiated carve-outs."),
        "Indemnity": ("Limit indemnity to defined third-party claims with defence and settlement control.", "Company's indemnity applies only to final third-party claims directly caused by the specified infringement or conduct. Company controls the defence and settlement, provided that no settlement imposes liability, admission, or non-monetary obligation on Client without Client's consent."),
        "Intellectual Property": ("Reserve reusable technology, know-how, and background materials.", "Company retains all rights in its Background IP, tools, methodologies, software, and know-how. Client receives only the licence expressly granted for its internal use of the Deliverables; no ownership or implied licence in Company Background IP transfers to Client."),
        "Data Protection / Security": ("Align operational obligations with documented instructions and realistic incident processes.", "Company will process Personal Data only on Client's documented instructions and will notify Client after confirming a Security Incident within [X] hours, taking into account the information reasonably available. Company's obligations are subject to Client's cooperation and do not require disclosure of other customers' confidential information."),
        "Audit / Compliance": ("Keep audits limited, secure, and non-disruptive.", "Client may audit no more than once in any [12]-month period, on [30] days' notice, during normal business hours, and at Client's cost, subject to confidentiality, security requirements, and use of independent assurance reports in lieu of on-site access where reasonably sufficient."),
        "Service Levels / Credits": ("Keep remedies finite and account for dependencies outside the provider's control.", "Service credits are Client's sole and exclusive monetary remedy for Service Level failures and are capped at [X%] of the monthly fees for the affected Service. Service Levels exclude failures caused by Client, third-party systems, scheduled maintenance, or events outside Company's reasonable control."),
        "Restrictive Covenants": ("Protect workforce stability without an unreasonable restraint.", "Any non-solicitation provision applies only to employees materially involved in the Services, excludes general advertisements and unsolicited applicants, and expires after [12] months; neither Party is liable for a hire made through a recruiter not instructed to target protected personnel."),
        "Dispute Resolution / Governing Law": ("Require a structured escalation process before costly litigation.", "Before commencing proceedings, each Party will escalate the dispute to senior business representatives for at least [15] days. Except for urgent equitable relief, neither Party may commence formal proceedings until that escalation process is complete."),
    },
}


def party_source_cue(source: str, lens: str) -> str:
    terms = ("customer", "client") if lens == "Client / Customer" else ("company", "service provider", "supplier", "vendor", "contractor")
    sentences = re.split(r"(?<=[.!?])\s+|\n+", source)
    cues = [sentence.strip() for sentence in sentences if any(re.search(r"\b" + re.escape(term) + r"\b", sentence, flags=re.I) for term in terms)]
    if cues:
        return "Direct source cue: " + " ".join(cues[:2])
    return "No explicit party reference was detected in this signal evidence; this lens is a category-based negotiation prompt, not an allocation finding."


def perspective_response(signal: dict, lens: str) -> dict[str, str]:
    objective, adjustment = SIGNAL_PARTY_OVERRIDES.get((lens, signal["title"]), PARTY_LENS.get(lens, {}).get(signal["category"], (
        "Clarify the obligation, exceptions, remedy, and allocation of cost and risk from this party's position.",
        "Add express wording identifying the responsible party, scope of obligation, exceptions, notice, cure process, remedy, and agreed financial limits.",
    )))
    return {"objective": objective, "adjustment": adjustment, "source_cue": party_source_cue(signal["source_text"], lens)}


def render_signal(signal: dict, lens: str | None = None) -> None:
    st.markdown(f"{badge(signal['severity'])} — **{signal['title']}**", unsafe_allow_html=True)
    st.write(signal["business_impact"])
    if lens:
        perspective = perspective_response(signal, lens)
        st.info(f"{lens} lens: {perspective['objective']}")
        st.caption(perspective["source_cue"])
    st.write(f"**Recommended action:** {signal['recommended_action']}")
    st.caption(f"Escalation: {signal['escalation']} · Clause {signal['clause_reference']} — {signal['clause_title']} · Page {signal['page_number'] or 'not available'}")
    with st.expander("Contract evidence"):
        st.code(signal["source_text"], language=None)
    st.divider()


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
                    "structure": structure_summary(clauses),
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
    structure = result.get("structure", {})
    if structure.get("derived", 0):
        st.warning(f"Document structure note: {structure['derived']} derived text fragment(s) were created because no reliable numbered or heading boundary was found. These are traceable review segments, not invented clause numbers.")
    else:
        st.caption(f"Structural extraction: {structure.get('numbered', 0)} numbered clause(s) and {structure.get('heading_derived', 0)} heading-derived clause(s).")
    memo_tab, signals_tab, drafting_tab, terms_tab, clauses_tab, human_tab = st.tabs(["Decision memo", "Deal signals", "Drafting responses", "Key terms", "Clauses", "Human review"])

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
        all_signals, client_signals, provider_signals = st.tabs(["All deal signals", "Client / Customer lens", "Company / Service Provider lens"])
        for tab, lens in ((all_signals, None), (client_signals, "Client / Customer"), (provider_signals, "Company / Service Provider")):
            with tab:
                if lens:
                    st.caption("Signals remain tied to the same agreement evidence. The lens supplies a distinct negotiation position and clearly marks where the source does not expressly name that party.")
                for signal in result["signals"]:
                    if category_filter and signal["category"] not in category_filter:
                        continue
                    if signal["severity"] not in severity_filter:
                        continue
                    render_signal(signal, lens)

    with drafting_tab:
        st.subheader("Party-segregated drafting responses")
        st.caption("Each proposal is linked to a red or amber signal. Choose the negotiating perspective you represent; sample language is not ready to sign and must be tailored by qualified counsel.")
        draftable = [signal for signal in result["signals"] if signal["severity"] in {"RED", "AMBER"}]
        client_drafts, provider_drafts = st.tabs(["Client / Customer proposals", "Company / Service Provider proposals"])
        for tab, lens in ((client_drafts, "Client / Customer"), (provider_drafts, "Company / Service Provider")):
            with tab:
                if not draftable:
                    st.info("No red or amber rule-triggered signals were found. The app cannot confirm that the agreement is risk-free.")
                for signal in draftable:
                    positioned = perspective_response(signal, lens)
                    with st.expander(f"{signal['severity']} — {signal['title']} · Clause {signal['clause_reference']} — {signal['clause_title']}"):
                        st.write(f"**{lens} drafting objective:** {positioned['objective']}")
                        st.caption(positioned["source_cue"])
                        st.write("**Position-specific sample clause:**")
                        st.code(positioned["adjustment"], language="text")
                        st.write("**Balanced baseline / cross-check:**")
                        st.code(drafting_response(signal)["clause"], language="text")
                        st.caption("Reconcile this proposal with the source evidence, the other party's position, governing law, and negotiated commercial thresholds before use.")
        if draftable:
            st.download_button("Download balanced drafting-response pack", drafting_pack(result), "deal_radar_drafting_responses.md", "text/markdown")

    with terms_tab:
        for term in result["terms"]:
            with st.expander(f"{term['label']}: {term['value']}"):
                st.caption(f"Clause {term['clause_reference']} — {term['clause_title']} · Page {term['page_number'] or 'not available'}")
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
