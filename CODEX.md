# Deal Contract Risk Radar

## Local-only design

This Streamlit app is a deterministic Python rule engine. It has no API key,
Ollama, LangChain, LangGraph, cloud model, or network dependency.

It parses text-based PDF and DOCX agreements, identifies numbered clauses,
runs commercial deal-risk rules against the entire extracted text, and links
each finding to the exact triggering contract language.

## Run in VS Code

1. Open this folder in VS Code and install Python 3.11+ with the Python extension.
2. Open **Terminal → New Terminal**.
3. Run:

   \`\`\`powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   .\run_contract_radar.ps1
   \`\`\`

The script creates a virtual environment, installs the three required Python
packages, and starts Streamlit. No environment variables are needed.

## What it screens

The Radar looks for commercially material rule patterns around:

- Change of control and assignment
- Termination and renewal
- Payment, pricing, MFN, and exclusivity
- Liability and indemnity
- IP, data/security, audit, and SLA exposure
- Restrictive covenants and dispute terms

## Limits

The tool is fast because it uses explicit rules rather than a language model.
It does not interpret ambiguous drafting, infer omitted provisions, or provide
legal advice. Every material decision must be checked by qualified counsel.
Scanned PDFs need OCR before they can be analysed.

