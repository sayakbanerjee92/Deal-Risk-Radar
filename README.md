# Deal-Risk-Radar
# Deal Contract Risk Radar

A fast, local-first Streamlit application that screens commercial contracts for deal-critical risk signals.

It is designed for early **PE/VC diligence**, **corporate development**, and **commercial approval** workflows. The app parses text-based PDF or DOCX agreements, identifies deterministic commercial signals, scores them, and links every finding to contract evidence.

> **Important:** This is a first-pass screening tool, not legal advice. It does not replace qualified legal review or determine whether to sign, invest, or proceed with a transaction.

## Why this exists

A full legal review is valuable but can be slow when an investment or commercial team first needs to know:

- Can the customer terminate after an acquisition?
- Does the agreement create revenue-retention or cash-collection risk?
- Are there exclusivity, MFN, pricing, liability, IP, data, or audit terms that need escalation?
- What needs confirmation before an investment committee or commercial approval decision?

The Radar provides an evidence-first answer without an API key, cloud model, Ollama, or network connection.

## Features

- Local PDF and DOCX text extraction
- Page-aware evidence for PDFs with selectable text
- Numbered-clause and preamble detection
- Deterministic screening of the whole extracted agreement
- Red / amber / green deal signals
- Transparent scorecard and decision memo
- Investment Diligence and Commercial Approval modes
- Human-review decisions and notes in the active Streamlit session
- Downloadable Markdown report
- No external data transfer

## Commercial risk areas

The current rule set screens for evidence relating to:

- Change of control and assignment
- Termination, renewal, and transition
- Payment terms and revenue commitments
- Pricing, MFN rights, rebates, and exclusivity
- Liability caps and uncapped exposure
- Indemnities
- Intellectual property and background-IP ownership
- Data protection and security obligations
- Audit and compliance rights
- Service levels and service credits
- Non-solicitation and other restrictions
- Dispute resolution and governing law

## Architecture

~~~text
PDF / DOCX upload
      ↓
Local Python text extraction
      ↓
Clause and preamble detection
      ↓
Deterministic commercial-risk rules
      ↓
Evidence-linked signals + transparent scorecard
      ↓
Decision memo + human review
~~~

The application uses Python rules rather than an LLM. This gives it predictable, fast local execution, but it cannot resolve ambiguous drafting or replace legal interpretation.

## Requirements

- Python 3.11+
- A text-based PDF or DOCX contract
- VS Code is recommended but not required

No API key, OpenAI account, Ollama installation, Docker installation, or internet connection is required after packages are installed.

## Installation

Open the project root in VS Code.

### Windows PowerShell

~~~powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\run_contract_radar.ps1
~~~

The launcher creates a virtual environment, installs dependencies, and starts Streamlit.

Alternatively:

~~~powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
~~~

Open the local Streamlit address shown in the terminal, normally http://localhost:8501.

## How to use

1. Choose Investment diligence or Commercial approval.
2. Upload a text-based PDF or DOCX agreement.
3. Select Run local deal radar.
4. Review the scorecard and priority signals.
5. Open evidence panels to inspect the literal triggering text.
6. Record a human decision: Proceed, Diligence further, Negotiate, Escalate, or Ignore.
7. Download the Markdown report if needed.

## Best contracts to test

Start with:

- Customer MSAs
- SaaS subscription agreements
- Enterprise commercial agreements
- Statements of Work
- Distribution, reseller, and strategic partnership agreements
- Data processing and security addenda
- IP licence agreements
- Key vendor agreements
- Agreements relevant to an acquisition, financing, or change of control

For PE/VC diligence, review top customer contracts by ARR or revenue first, followed by agreements with upcoming renewals, exclusivity, unusual liability, IP ownership, or assignment restrictions.

## Limitations

- Scanned PDFs must be OCR'd before analysis.
- Rules identify text patterns; they do not interpret jurisdiction, negotiation history, or business intent.
- A match is a review prompt, not a legal conclusion.
- A missing signal does not prove that no risk exists.
- Human decisions are session-only and are not stored in a database.
- Do not upload confidential contracts to a public GitHub repository.

## Project structure

~~~text
.
├── app.py                    # Streamlit app and deterministic rule engine
├── run_contract_radar.ps1    # Windows setup and launcher
├── requirements.txt          # Minimal Python dependencies
├── CODEX.md                  # Local-development notes
└── README.md                 # This guide
~~~

## Publishing to GitHub

GitHub stores files and commits; it does not automatically upload Codex or ChatGPT conversations.

### 1. Create an empty repository

On GitHub, select **New repository**, name it (for example, deal-contract-risk-radar), select **Private** for contract-related development, and do not initialize it with a README.

### 2. Initialize and commit this project

Run these commands in the VS Code PowerShell terminal from this folder:

~~~powershell
git init
git add app.py requirements.txt run_contract_radar.ps1 README.md CODEX.md
git commit -m "Initial Deal Contract Risk Radar"
git branch -M main
~~~

### 3. Connect and push

Replace YOUR-USERNAME and the repository name:

~~~powershell
git remote add origin https://github.com/YOUR-USERNAME/deal-contract-risk-radar.git
git push -u origin main
~~~

GitHub will ask you to authenticate in a browser or through Git Credential Manager.

## Preserving the prior chat

Do not commit secrets, contract text, API keys, or personal information.

If you want to preserve the design discussion:

1. In Codex, copy or export only the relevant messages.
2. Create docs/development-notes.md.
3. Add a cleaned, non-confidential summary of goals, architecture decisions, and limitations.
4. Commit it deliberately:

~~~powershell
git add docs/development-notes.md
git commit -m "Document product and architecture decisions"
git push
~~~

For a public repository, a concise decision record is better than publishing the full chat transcript.


