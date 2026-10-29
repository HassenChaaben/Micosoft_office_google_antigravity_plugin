"""
Gemini Streaming AI Engine
Streams reasoning, analysis, and document transformations using Gemini 3.8 Flash High, Pro, and Deep Think.
Supports live Google Gemini API streaming (via API Key or Subscription Bridge) and intelligent local domain synthesis.
"""

import os
import re
import time
import json
import asyncio
from typing import AsyncGenerator, Dict, Any
from . import antigravity_backend
from .auth import auth_bridge
from .skills_library import get_skill_definition

class GeminiClient:
    def __init__(self):
        pass

    async def stream_chat(self, payload: Dict[str, Any]) -> AsyncGenerator[str, None]:
        prompt = payload.get("prompt", "")
        model = payload.get("model", "gemini-3.8-flash-high")
        host = payload.get("host", "Word")
        web_search = payload.get("web_search", True)
        doc_context = payload.get("document_context", "")

        backend = os.environ.get("GEMINI_OFFICE_BACKEND", "auto").lower()
        errors = []

        # 0. Antigravity subscription via the signed-in `agy` CLI
        agy_path = antigravity_backend.find_agy() if backend != "offline" else None
        if agy_path:
            events = antigravity_backend.stream_events(agy_path, prompt, host, doc_context, model)
            has_streamed = False
            while True:
                event = await asyncio.to_thread(next, events, None)
                if event is None:
                    break
                kind, text = event
                if kind == "delta":
                    has_streamed = True
                    yield f"data: {json.dumps({'delta': text})}\n\n"
                elif kind == "error":
                    errors.append(text)
                    if has_streamed:
                        yield f"data: {json.dumps({'delta': chr(10) * 2 + '> ⚠️ ' + text})}\n\n"
                    break
                else:
                    break
            if has_streamed:
                yield "data: [DONE]\n\n"
                return
        elif backend != "offline":
            errors.append("Antigravity CLI (`agy`) was not found on this PC. Install/sign in to Antigravity CLI.")

        api_key = auth_bridge.get_api_key() if backend != "offline" else None

        # 1. Live Google Gemini API Stream (if API Key configured)
        if api_key:
            try:
                import google.generativeai as genai
                genai.configure(api_key=api_key)

                # Map UI model selector to Google Gemini model name
                model_map = {
                    "gemini-3.8-flash-high": "gemini-flash-latest",
                    "gemini-pro": "gemini-pro-latest",
                    "gemini-deep-think": "gemini-pro-latest"
                }
                model_id = model_map.get(model, "gemini-flash-latest")

                system_instruction = (
                    f"You are Gemini Antigravity, an autonomous executive AI operating inside Microsoft {host}. "
                    "Generate comprehensive, structured, publication-grade deliverables with headings, tables, bullet points, and actionable details. "
                    "Do not include conversational filler like 'Sure, here is...' or 'I hope this helps'."
                )

                full_prompt = prompt
                if doc_context:
                    full_prompt = (
                        f"Active Microsoft {host} Document Context:\n"
                        f"\"\"\"\n{doc_context}\n\"\"\"\n\n"
                        f"User Request:\n{prompt}"
                    )

                gen_model = genai.GenerativeModel(model_id, system_instruction=system_instruction)
                response = gen_model.generate_content(full_prompt, stream=True)

                has_streamed = False
                for chunk in response:
                    if hasattr(chunk, "text") and chunk.text:
                        has_streamed = True
                        sse_data = json.dumps({"delta": chunk.text})
                        yield f"data: {sse_data}\n\n"

                if has_streamed:
                    yield "data: [DONE]\n\n"
                    return
            except Exception as e:
                print(f"[GeminiClient] Live Google API call failed: {e}")
                errors.append(f"Gemini API key call failed: {e}")

        if backend != "offline":
            message = (
                "### ⚠️ No AI answer\n\n"
                "The request could not reach a real AI model, so nothing was generated.\n\n"
                + "\n".join(f"- {err}" for err in errors)
                + "\n\nCheck that Antigravity CLI is signed in (run `agy models` in a terminal), then try again."
            )
            yield f"data: {json.dumps({'delta': message})}\n\n"
            yield "data: [DONE]\n\n"
            return

        # 2. Offline template mode (only when GEMINI_OFFICE_BACKEND=offline, e.g. tests/demos)
        response_text = self._generate_intelligent_response(prompt, host, doc_context, web_search, model)

        # Stream words as SSE deltas to simulate realistic real-time token streaming
        words = response_text.split(" ")
        is_test = bool(os.environ.get("PYTEST_CURRENT_TEST"))
        for i, word in enumerate(words):
            chunk = word + (" " if i < len(words) - 1 else "")
            sse_data = json.dumps({"delta": chunk})
            yield f"data: {sse_data}\n\n"
            if not is_test:
                await asyncio.sleep(0.012)

        yield "data: [DONE]\n\n"

    def _generate_intelligent_response(self, prompt: str, host: str, doc_context: str, web_search: bool, model: str) -> str:
        clean_p = prompt.strip()
        lower_p = clean_p.lower()

        # Skill 1: /check-doc
        if lower_p.startswith("/check-doc"):
            return (
                "### Document Review & Tone Audit\n\n"
                "**1. Executive Clarity & Structure:**\n"
                "- The central premise is well-defined, but sections 2 and 4 suffer from passive voice and unnecessary nominalizations.\n"
                "- **Paragraph 2 flow:** Transition from technical constraints to commercial outcomes could be strengthened.\n\n"
                "**2. Tone & Consistency:**\n"
                "- Professional and direct. Replaced 4 instances of weak modal verbs (*'might consider'* → *'recommends'*).\n\n"
                "**3. Suggested Replacement for Selection:**\n\n"
                "> *\"The proposed Gemini Antigravity architecture deploys a native localhost daemon, establishing zero-latency COM automation and direct subscription authentication for Microsoft Word and PowerPoint.\"*\n\n"
                "*(Click 'Apply to Word Selection' below to update your document instantly)*"
            )

        # Skill 2: /copy-edit
        if lower_p.startswith("/copy-edit"):
            return (
                "### Copy-Edit & Grammar Polish\n\n"
                "**Corrections Applied:**\n"
                "1. **Conciseness:** Eliminated redundant verbiage (*'at the present point in time'* → *'currently'*; *'due to the fact that'* → *'because'*).\n"
                "2. **Punctuation:** Enforced Oxford commas across all lists.\n"
                "3. **Active Voice:** Converted 3 passive clauses to direct active phrasing.\n\n"
                "**Polished Text:**\n\n"
                "> *\"Gemini Antigravity streamlines document review, flags contract liabilities, and generates executive slide decks autonomously inside Microsoft Office.\"*\n\n"
                "*(Click 'Apply to Word Selection' to insert the revised version)*"
            )

        # Skill 3: /summarize-contract
        if lower_p.startswith("/summarize-contract"):
            return (
                "### Executive Contract Summary\n\n"
                "| Term | Details |\n"
                "| :--- | :--- |\n"
                "| **Effective Date** | October 28, 2026 |\n"
                "| **Parties** | Cloud4Job / Enterprise Partner |\n"
                "| **Deliverables** | Gemini Antigravity Office Plugin deployment across Word and PowerPoint |\n"
                "| **Governing Law** | Italian Jurisdiction / EU GDPR compliant |\n"
                "| **Liability Cap** | Capped at 100% of aggregate fees paid in past 12 months |\n\n"
                "**Key Obligations & Notice Windows:**\n"
                "- 30 days written notice required for termination without cause.\n"
                "- 14-day cure period for material breach."
            )

        # Skill 4: /flag-issues
        if lower_p.startswith("/flag-issues"):
            return (
                "### Flagged Contract Risks & Ambiguities\n\n"
                "⚠️ **High Risk — Uncapped Indemnity (Clause 5.1):**\n"
                "- The indemnity clause holds the supplier liable for indirect and consequential damages with no aggregate dollar ceiling.\n"
                "- *Mitigation:* Insert mutual liability cap of 12 months' fees.\n\n"
                "⚠️ **Moderate Risk — Ambiguous Milestone Acceptance (Clause 3.4):**\n"
                "- Client retains subjective approval without defined verification test criteria.\n"
                "- *Mitigation:* Add deemed acceptance within 10 business days of delivery.\n\n"
                "*(Click 'Insert Review Comment' below to place inline annotations directly on the Word text)*"
            )

        # Skill 5: /industry-overview
        if lower_p.startswith("/industry-overview"):
            grounding_note = " *(Grounded with live Google Web Search)*" if web_search else ""
            return (
                f"### Industry & Market Overview: Enterprise Office AI Agents{grounding_note}\n\n"
                "**1. Market Dynamics & Growth:**\n"
                "- The enterprise AI productivity market is projected to reach $85B by 2028, growing at a 32.4% CAGR.\n"
                "- Key shift: Movement from disconnected web chatbots toward **ambient, embedded copilots** with deep document awareness.\n\n"
                "**2. Competitive Matrix:**\n"
                "- **Microsoft Copilot:** Deep Office 365 integration, but limited local MCP tool extensibility.\n"
                "- **Claude Office:** Smooth UI with skills, but requires specific enterprise seat tiers.\n"
                "- **Gemini Antigravity:** Zero-API-key Google subscription bridge, local MCP connectivity, and hybrid COM automation."
            )

        # Skill 6: /storylining
        if lower_p.startswith("/storylining"):
            return (
                "### PowerPoint Executive Presentation Storyline\n\n"
                "**Thesis:** *Autonomous Office AI transforms desktop workflows by embedding directly into user documents with zero friction.*\n\n"
                "**Slide 1: Executive Overview**\n"
                "- Problem: Knowledge workers lose 40% of their workday jumping between web chatbots and Office documents.\n"
                "- Solution: Gemini Antigravity brings autonomous AI directly into the Word and PowerPoint ribbon.\n\n"
                "**Slide 2: System Architecture**\n"
                "- Native Ribbon button loads WebView2 Taskpane.\n"
                "- Localhost daemon eliminates CORS and manages Google subscription auth.\n"
                "- Dual-engine COM + Office.js executes document edits in real time.\n\n"
                "**Slide 3: Enterprise Value & Velocity**\n"
                "- 3.5x faster contract review and presentation authoring.\n"
                "- Seamless connectivity to Notion, GitHub, and local MCP skills.\n\n"
                "*(Click 'Create PowerPoint Slide' below to build this deck directly into your active presentation)*"
            )

        # Excel Skill 1: /analyze-sheet
        if lower_p.startswith("/analyze-sheet"):
            return (
                "### Excel Spreadsheet Analysis & Insights\n\n"
                "**1. Dataset Health & Summary Statistics:**\n"
                "- Analyzed active used range across columns. No null key identifiers or circular formula references detected.\n"
                "- **Outlier Detection:** Row 4 exhibits a +34.2% sequential deviation driven by seasonal surge.\n\n"
                "**2. Trend Trajectory:**\n"
                "- Annualized compound growth rate (CAGR) stands at **41.8%**.\n"
                "- Gross Margin averages **76.4%**, in line with top-quartile SaaS metrics.\n\n"
                "*(Click 'Insert Financial Model' or 'Generate Chart' to apply updates directly into Excel)*"
            )

        # Excel Skill 2: /financial-model
        if lower_p.startswith("/financial-model"):
            return (
                "### Autonomous 3-Statement Financial Model\n\n"
                "| Metric | FY2024 | FY2025 | FY2026E | FY2027E | CAGR |\n"
                "| :--- | :--- | :--- | :--- | :--- | :--- |\n"
                "| **Subscription Revenue** | $5,200,000 | $8,100,000 | $12,800,000 | $19,500,000 | 55.4% |\n"
                "| **Professional Services** | $800,000 | $1,100,000 | $1,400,000 | $1,800,000 | 31.0% |\n"
                "| **Total Revenue** | $6,000,000 | $9,200,000 | $14,200,000 | $21,300,000 | 52.6% |\n"
                "| **Gross Profit** | $4,800,000 | $7,400,000 | $11,600,000 | $17,600,000 | 54.1% |\n"
                "| **EBITDA** | $1,200,000 | $2,100,000 | $3,800,000 | $6,200,000 | 72.8% |\n"
                "| **EBITDA Margin** | 20.0% | 22.8% | 26.7% | 29.1% | - |\n\n"
                "*(Click 'Insert Financial Model into Excel' to write live formulas and formatted tables directly into your spreadsheet)*"
            )

        # Excel Skill 3: /formula-helper
        if lower_p.startswith("/formula-helper"):
            return (
                "### Excel Formula Recommendations\n\n"
                "**1. Dynamic Lookup (XLOOKUP):**\n"
                "```excel\n"
                "=XLOOKUP(A2, 'ReferenceData'!A:A, 'ReferenceData'!E:E, \"Not Found\", 0)\n"
                "```\n"
                "- *Advantage:* Replaces fragile VLOOKUP without hardcoded column indices.\n\n"
                "**2. CAGR Multi-Year Calculation:**\n"
                "```excel\n"
                "=(End_Value / Start_Value) ^ (1 / Number_Of_Years) - 1\n"
                "```\n\n"
                "**3. Dynamic Array Unique Filter:**\n"
                "```excel\n"
                "=SORT(UNIQUE(FILTER(A2:D100, C2:C100 > 10000)))\n"
                "```"
            )

        # Outlook Skill 1: /reply-email
        if lower_p.startswith("/reply-email"):
            return (
                "### Outlook Executive Email Reply Draft\n\n"
                "**Subject:** Re: Q4 Enterprise Strategy & Autonomous Office AI Deliverables\n\n"
                "> Hi Ergys,\n>\n"
                "> Thank you for the follow-up. We have successfully unified the Gemini Antigravity platform across all four Microsoft Office applications: **Word, Excel, PowerPoint, and Outlook**.\n>\n"
                "> **Key Progress Delivered:**\n"
                "> 1. **Excel Engine:** Autonomous reading, formula writing (`XLOOKUP`, `CAGR`), financial models, and native chart insertion.\n"
                "> 2. **Outlook Engine:** Thread context analysis, formatted HTML reply drafting, and one-click calendar sync.\n"
                "> 3. **Cross-App Pipelines:** Automated `/doc-to-slides` (Word to PowerPoint) and `/excel-to-report` (Excel to Word).\n>\n"
                "> I've scheduled a 30-minute verification review for tomorrow at 10:00 AM.\n>\n"
                "> Best regards,\n> **Hassen**\n\n"
                "*(Click 'Draft Reply in Outlook' below to open this formatted message in Outlook)*"
            )

        # Outlook Skill 2: /summarize-thread
        if lower_p.startswith("/summarize-thread"):
            return (
                "### Email Thread Synthesis & Action Items\n\n"
                "**Core Subject:** Enterprise License Rollout & Security Certification\n\n"
                "**Executive Summary:**\n"
                "- The client confirmed budget approval for 250 enterprise seats pending SOC2 and GDPR verification.\n"
                "- Delivery window is targeted for November 15, 2026.\n\n"
                "**Action Items:**\n"
                "- [ ] **Security:** Forward local daemon localhost SSL specification by Thursday.\n"
                "- [ ] **Legal:** Update Clause 5.1 mutual indemnity ceiling in the Word master agreement.\n"
                "- [ ] **Deployment:** Distribute standalone `GeminiAntigravityOffice_Setup.exe` to IT administrators."
            )

        # Cross-App Pipeline: /doc-to-slides
        if lower_p.startswith("/doc-to-slides"):
            return (
                "### Cross-App Pipeline: Word Document to PowerPoint Deck\n\n"
                "**Source:** Active Microsoft Word Document  \n"
                "**Target:** Microsoft PowerPoint Executive Presentation  \n\n"
                "- **Slide 1:** Executive Summary & Strategic Premise\n"
                "- **Slide 2:** Current Bottlenecks & Operational Constraints\n"
                "- **Slide 3:** Unified 4-App System Architecture (Word, Excel, PPT, Outlook)\n"
                "- **Slide 4:** Quantified Business Velocity & Enterprise ROI\n\n"
                "*(Click 'Build PowerPoint Deck' below to autonomously generate these slides in PowerPoint)*"
            )

        # Cross-App Pipeline: /excel-to-report
        if lower_p.startswith("/excel-to-report"):
            return (
                "### Cross-App Pipeline: Excel Table to Word Executive Report\n\n"
                "**Source:** Active Microsoft Excel Worksheet  \n"
                "**Target:** Microsoft Word Publication-Grade Briefing  \n\n"
                "Extracted 5 quarters of revenue and EBITDA performance. "
                "Synthesized an executive narrative with formatted data tables, margin CAGR formulas, and strategic recommendations.\n\n"
                "*(Click 'Insert Formatted Report into Word' to generate the document)*"
            )

        # Connectors: /notion
        if lower_p.startswith("/notion"):
            query = clean_p.replace("/notion", "").strip() or "workspace recent pages"
            return (
                f"### Notion Workspace Connected\n\n"
                f"**Query:** `{query}`\n\n"
                "**Synced Pages & Databases:**\n"
                "- **Robot-Voice Hub** (`https://www.notion.so/3ffa15fdc7358072a927fb1bfb6ef6ab`)\n"
                "- **Tasks Tracker** (Database ID: `338a15fdc735806198bbe487801faa93`) — Status: *In Progress*\n"
                "- **Memories Hub** (`https://www.notion.so/405a15fdc7358113a520ccf527ade587`) — 24 items synced\n\n"
                "*(You can reference any Notion page or database properties directly in your document)*"
            )

        # Specialized Deep Synthesis: Quantum Computing Report
        if any(w in lower_p for w in ["quantum", "quantuim", "qubit", "shor"]):
            return (
                "# Executive Report: State and Enterprise Horizon of Quantum Computing\n\n"
                "**Prepared by:** Gemini Antigravity Office  \n"
                "**Target Platform:** Microsoft Word Briefing  \n\n"
                "---\n\n"
                "## 1. Executive Summary\n"
                "Quantum computing represents an architectural disruption in computational complexity. "
                "By moving beyond classical binary transistors (`0` or `1`) to quantum mechanical states of **superposition**, **entanglement**, and **constructive interference**, quantum processors solve previously intractable computational problems in cryptography, combinatorial optimization, and molecular simulation.\n\n"
                "---\n\n"
                "## 2. Core Scientific Principles\n"
                "- **Superposition:** An $N$-qubit state exists in an arbitrary linear combination of all $2^N$ computational basis states simultaneously: $|\\psi\\rangle = \\sum_{x} \\alpha_x |x\\rangle$, allowing simultaneous evaluation across exponential parameter spaces.\n"
                "- **Quantum Entanglement:** Quantum correlations enable non-local state linkages where measuring one qubit instantaneously determines the state of its entangled partner, forming the foundation for quantum error correction.\n"
                "- **Phase Estimation & Interference:** Quantum algorithms design interference patterns where correct algorithmic trajectories reinforce constructively, while erroneous trajectories cancel destructively.\n\n"
                "---\n\n"
                "## 3. Hardware Modalities & Industry Benchmarks\n\n"
                "| Architecture | Leading Pioneers | Coherence Time | Primary Engineering Bottleneck |\n"
                "| :--- | :--- | :--- | :--- |\n"
                "| **Superconducting Transmons** | IBM Quantum (Heron 133Q), Google Quantum AI | ~100 µs | High thermal noise, requires 15 mK dilution refrigeration |\n"
                "| **Trapped Ions** | IonQ (Forte), Quantinuum (H2) | Hours | Slower two-qubit gate execution times (~1–10 µs) |\n"
                "| **Neutral Atoms (Rydberg)** | QuEra (Aquila), Pasqal | ~1–5 s | Laser phase noise and shuttle transport overhead |\n"
                "| **Photonic Quantum** | PsiQuantum, Xanadu | Room-temp | Non-deterministic single-photon sources and switch loss |\n\n"
                "---\n\n"
                "## 4. Key Enterprise Application Verticals\n"
                "1. **Post-Quantum Cryptography (PQC) & Security:**\n"
                "   - Shor's algorithm renders existing RSA-2048 and ECC public key infrastructure vulnerable.\n"
                "   - NIST has formalized post-quantum standards: **FIPS 203 (ML-KEM)**, **FIPS 204 (ML-DSA)**, and **FIPS 205 (SLH-DSA)**.\n"
                "2. **Chemical & Material Science Simulation:**\n"
                "   - Variational Quantum Eigensolvers (VQE) calculate exact ground-state electron orbitals for battery cathode design, nitrogen fixation, and polymer catalysts.\n"
                "3. **Logistics & Portfolio Optimization:**\n"
                "   - Quantum Approximate Optimization Algorithm (QAOA) resolves quadratic unconstrained binary optimization (QUBO) problems for global supply chain scheduling and financial risk hedging.\n\n"
                "---\n\n"
                "## 5. Strategic Roadmap & Recommendations (2026–2030)\n"
                "1. **Enact Crypto-Agility Audits:** Inventory all enterprise asymmetric keys and transition to hybrid classical-post-quantum algorithms.\n"
                "2. **Implement Cloud Quantum Hybrid Workflows:** Connect classical HPC clusters with cloud-accessible quantum QPUs via hybrid frameworks (Qiskit, Cirq, PennyLane).\n"
                "3. **Focus on Fault-Tolerant Logical Qubits:** Monitor error-corrected logical qubit thresholds rather than raw physical qubit counts."
            )

        # Intent 1: Daily Tasks / Task Tracker / To-Do Table
        if any(w in lower_p for w in ["task", "todo", "to-do", "tracker", "tracking", "daily", "backlog", "kanban"]):
            return (
                "# Daily Tasks Tracking Matrix\n\n"
                "**Prepared by:** Gemini Antigravity Office  \n"
                "**Date:** October 28, 2026  \n"
                "**Workflow Mode:** Executive Task Management  \n\n"
                "---\n\n"
                "### 1. Active Task Matrix\n\n"
                "| Task ID | Task Description | Category | Priority | Assigned To | Due Date | Status | Deliverable / Output |\n"
                "| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |\n"
                "| **TSK-101** | Finalize Q4 enterprise deliverables and milestones | Strategy | 🔴 High | Hassen | 2026-10-29 | In Progress | Aligning deliverables with client SLA |\n"
                "| **TSK-102** | Audit Microsoft Office plugin document insertion | Engineering | 🔴 High | Engineering | 2026-10-28 | Completed | Verified multi-strategy Word insertion |\n"
                "| **TSK-103** | Prepare executive financial summary and graphs | Finance | 🟡 Medium | Financial Analyst | 2026-10-30 | In Progress | 3-statement model & revenue CAGR |\n"
                "| **TSK-104** | Client demonstration and onboarding session | Client Success | 🟢 Normal | Hassen | 2026-10-31 | Scheduled | Share 1-click installer and setup guide |\n"
                "| **TSK-105** | Daily cross-functional standup and sync | Operations | 🟢 Normal | Team | 2026-10-29 | Planned | 10:00 AM Teams review meeting |\n\n"
                "---\n\n"
                "### 2. Operational Summary & Key Performance Indicators\n"
                "- **Total Tasks:** 5 | **High Priority:** 2 | **Completed:** 1 | **In Progress:** 2 | **Scheduled:** 1\n"
                "- **Today's Core Focus:** Complete the automated document insertion validation and client walk-through.\n"
                "- **Next Milestone:** Production deployment and client sign-off."
            )

        # Intent 2: Project Roadmap / Milestone Schedule Table
        if any(w in lower_p for w in ["roadmap", "timeline", "milestone", "project plan", "schedule"]):
            return (
                "# Project Roadmap & Milestone Schedule\n\n"
                "**Project:** Enterprise Office AI Suite Deployment  \n"
                "**Phase:** Production Rollout  \n\n"
                "---\n\n"
                "| Phase | Milestone Name | Key Deliverables | Target Date | Owner | Status |\n"
                "| :--- | :--- | :--- | :--- | :--- | :--- |\n"
                "| **Phase 1** | Architecture & Engines | Word, Excel, PowerPoint & Outlook DAEMON engines | 2026-10-20 | Architecture Lead | Completed |\n"
                "| **Phase 2** | UI & Taskpane Experience | Pixel-perfect Claude replica with slash popovers | 2026-10-26 | Frontend Lead | Completed |\n"
                "| **Phase 3** | Automated Client Setup | Standalone 1-click Windows installer with silent TLS | 2026-10-28 | DevOps Lead | Completed |\n"
                "| **Phase 4** | Client Validation & Sign-off | End-to-end user verification and feature testing | 2026-10-30 | Hassen | In Progress |\n"
                "| **Phase 5** | Production Rollout | Deployment across all client workstations | 2026-11-05 | Enterprise Team | Planned |\n\n"
                "---\n\n"
                "### Execution Invariants:\n"
                "1. Zero manual certificate installation required for end clients.\n"
                "2. 100% reliable document insertion across Word, Excel, and PowerPoint."
            )

        # Intent 3: General Table Creation Request
        if "table" in lower_p or "matrix" in lower_p or "grid" in lower_p:
            subj_match = re.search(r"(?:table|matrix|grid)\s+(?:of|for|about|showing|comparing)?\s*([a-zA-Z0-9\s_-]+)", clean_p, re.IGNORECASE)
            subject = subj_match.group(1).strip().title() if subj_match else "Operational Tracking"
            return (
                f"# Formatted Table: {subject}\n\n"
                f"**Generated by:** Gemini Antigravity Office  \n"
                f"**Target Host:** Microsoft {host}  \n\n"
                f"---\n\n"
                f"| Item # | Item Description | Category | Metric / Specification | Assigned Lead | Target Date | Status |\n"
                f"| :--- | :--- | :--- | :--- | :--- | :--- | :--- |\n"
                f"| **01** | Primary Deliverable: {subject} Core Baseline | Core Operations | Grade A Enterprise Spec | Project Lead | 2026-10-29 | In Progress |\n"
                f"| **02** | Security & Compliance Audit | Governance | EU GDPR & Local Storage | Security Team | 2026-10-30 | Approved |\n"
                f"| **03** | Performance Optimization & Latency Check | Engineering | Sub-10ms response time | Technical Lead | 2026-10-28 | Completed |\n"
                f"| **04** | Client Demonstration & User Acceptance | Client Success | 100% Feature Validation | Hassen | 2026-10-31 | Scheduled |\n"
                f"| **05** | Final Review & Production Release | Leadership | Executive Sign-off | Steering Group | 2026-11-02 | Planned |\n\n"
                f"---\n\n"
                f"*(Click 'Insert into Word Document' below to place this formatted table directly into your document)*"
            )

        # Intent 4: Meeting Notes & Action Items
        if any(w in lower_p for w in ["meeting", "minutes", "agenda", "sync", "standup"]):
            return (
                "# Executive Meeting Minutes & Action Items\n\n"
                "**Meeting Subject:** Enterprise Autonomous Office Project Sync  \n"
                "**Date & Time:** October 28, 2026 | 10:00 AM - 10:45 AM  \n"
                "**Attendees:** Hassen Ben Chaaben, Engineering Lead, Enterprise Client Team  \n\n"
                "---\n\n"
                "### 1. Agenda Topics Discussed:\n"
                "1. Verification of document insertion into Microsoft Word.\n"
                "2. Standalone client installer and silent localhost SSL trust.\n"
                "3. Cross-application orchestration (/doc-to-slides and /excel-to-report).\n\n"
                "---\n\n"
                "### 2. Action Items & Owners:\n\n"
                "| Item # | Action Description | Owner | Due Date | Priority | Status |\n"
                "| :--- | :--- | :--- | :--- | :--- | :--- |\n"
                "| **ACT-1** | Deploy multi-tier Word HTML insertion fallback | Engineering | 2026-10-28 | 🔴 High | Completed |\n"
                "| **ACT-2** | Verify taskpane chat table rendering | UI/UX Lead | 2026-10-28 | 🔴 High | Completed |\n"
                "| **ACT-3** | Deliver 1-click setup package and user guide | Hassen | 2026-10-30 | 🟡 Medium | In Progress |\n"
                "| **ACT-4** | Schedule final executive sign-off session | Client Success | 2026-10-31 | 🟢 Normal | Scheduled |\n"
            )

        # Intent 5: Budget / Expenses / Pricing Table
        if any(w in lower_p for w in ["budget", "expense", "cost", "pricing", "financial plan"]):
            return (
                "# Enterprise Budget & Expenditure Breakdown\n\n"
                "**Fiscal Year:** FY2026 - Q4 Operational Budget  \n"
                "**Currency:** USD ($)  \n\n"
                "---\n\n"
                "| Cost Center | Description | Budgeted ($) | Actual ($) | Variance ($) | Variance % | Status |\n"
                "| :--- | :--- | :--- | :--- | :--- | :--- | :--- |\n"
                "| **Engineering** | Cloud Infrastructure & AI Compute | $150,000 | $142,500 | -$7,500 | -5.0% | On Track |\n"
                "| **Product Development** | Client Integration & Desktop Testing | $95,000 | $91,200 | -$3,800 | -4.0% | On Track |\n"
                "| **Quality Assurance** | Automated Test Suites & Security Audit | $45,000 | $45,000 | $0 | 0.0% | Balanced |\n"
                "| **Client Success** | Training, Documentation & Onboarding | $35,000 | $32,100 | -$2,900 | -8.3% | Favorable |\n"
                "| **Operations** | Licenses, Tooling & Compliance Certs | $25,000 | $26,500 | +$1,500 | +6.0% | Monitored |\n"
                "| **TOTAL** | **Consolidated Q4 Operating Expenditures** | **$350,000** | **$337,300** | **-$12,700** | **-3.6%** | **Favorable** |\n\n"
                "---\n\n"
                "### Financial Commentary:\n"
                "- Total expenditures are currently **$12,700 (-3.6%) below budget**, demonstrating disciplined capital management.\n"
                "- Cost of compute optimized through native local daemon caching and local synthesis."
            )

        # General Document/Report Synthesis
        topic_match = re.search(r"(?:report|write|create|summary|about|on)\s+(?:of|on|about|a)?\s*([a-zA-Z0-9\s_-]+)", clean_p, re.IGNORECASE)
        topic = topic_match.group(1).strip().title() if topic_match else "Strategic Analysis & Executive Overview"
        
        return (
            f"# Executive Report: {topic}\n\n"
            f"**Prepared by:** Gemini Antigravity Office  \n"
            f"**Environment:** Microsoft {host}  \n"
            f"**Model:** `{model}`  \n\n"
            f"---\n\n"
            f"## 1. Executive Summary\n"
            f"This comprehensive briefing delivers a detailed operational and strategic analysis of **{topic}**. "
            f"The objective is to synthesize current market dynamics, technological benchmarks, operational risk vectors, and actionable implementation priorities.\n\n"
            f"---\n\n"
            f"## 2. Key Industry & Technical Drivers\n"
            f"- **Technological Innovation:** Rapid maturation of foundational models and autonomous tooling is accelerating enterprise adoption.\n"
            f"- **Efficiency Gains:** Organizations adopting structured automation pipelines report a 35% reduction in cycle times.\n"
            f"- **Data Architecture:** Success hinges on unified, ambient context awareness across core enterprise documents and databases.\n\n"
            f"---\n\n"
            f"## 3. Comparative Architecture Matrix\n\n"
            f"| Dimension | Traditional Approach | Modern Automated Standard | Impact Delta |\n"
            f"| :--- | :--- | :--- | :--- |\n"
            f"| **Execution Velocity** | Manual / Disconnected | Ambient / Embedded Office Agent | 4.2x faster turn-around |\n"
            f"| **Error Rate** | Prone to copy-paste drift | Direct COM / In-Place API | 99.4% context fidelity |\n"
            f"| **Security & Privacy** | Web-portal exposure | Localhost Daemon Bridge | Zero public endpoint leakage |\n\n"
            f"---\n\n"
            f"## 4. Operational Recommendations & Next Steps\n"
            f"1. **Immediate Integration:** Embed directly into existing Word and PowerPoint workflows.\n"
            f"2. **Continuous Governance:** Implement audit trails and validation gates for all automated modifications.\n"
            f"3. **Knowledge Base Sync:** Connect active workspaces (Notion, GitHub, local files) to maintain dynamic grounding.\n"
        )

gemini_client = GeminiClient()
