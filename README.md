# WSQ — Applications Integration with Power Apps and Power Automate

Spot, test, build, join and repair application integrations on the Microsoft Power Platform through 14 progressive hands-on labs.

| Course detail | Information |
|---|---|
| Course code | `TGS-2022015539` |
| Programme | WSQ |
| TSC | Applications Integration · `ICT-DIT-3003-1.1`, Proficiency Level 3 |
| Duration | 2 days · 16 hours (13.5 hours delivery + 2.5 hours assessment) |
| Registration | [View course details and register](https://www.tertiarycourses.com.sg/wsq-applications-integration-with-power-apps-and-power-automate.html) |
| Funding | Up to 70% course fee funding (50% baseline for SG/PR aged 21+; 70% for Singapore Citizens aged 40+ under MCES and SME-sponsored staff). Eligibility and terms apply — see the course page. |
| Courseware version | v8.0 |

---

## About the course

Most business software is bought, not built — the value you add is making the pieces talk
to each other. This course teaches application integration with the Microsoft Power
Platform: how to spot an integration opportunity, test whether it is feasible, build the
automation in Power Automate, build the interface in Power Apps, join the two together,
and then diagnose and repair the integration when it breaks.

## Learning outcomes

| | Outcome |
|---|---|
| **LO1** | Identify opportunities to use Power Platform apps and perform a feasibility scan for potential application integration |
| **LO2** | Utilise and test Power Automate flows for business automation |
| **LO3** | Create Power Apps canvas apps and verify their functionalities |
| **LO4** | Improve Power Apps with Power Automate flows, and resolve integration issues |

## Topics covered

1. Opportunities for using Power Platform apps and feasibility scanning
2. Power Automate — automated, scheduled and approval flows with conditions and branching
3. Power Apps — canvas apps from data and from blank, custom connectors and cross-platform verification
4. Integrating Power Apps and Power Automate — calling flows, returning data, end-to-end approvals and repairing broken integrations

## Labs

14 hands-on labs, each building on the artefact the previous one produced. Work them in
order — skipping ahead leaves you without the flow or the table the next lab expects.

| Topic | Labs | What you build |
|---|---|---|
| 1 · Opportunities for Using Power Platform Apps | 1–2 | An Integration Opportunity Register and a six-criteria Feasibility Assessment Matrix |
| 2 · Power Automate | 3–6 | Automated, scheduled and approval flows with conditions and branching |
| 3 · Power Apps | 7–10 | Canvas apps from data and from blank, a custom connector over a REST API, cross-platform verification |
| 4 · Integrate Power Apps and Power Automate | 11–14 | An app that calls a flow, returns data, runs an end-to-end approval — then breaks, diagnoses and repairs it |

Full index and instructions: **[labs/README.md](labs/README.md)**

Each lab folder holds its lab sheet, the mock-data workbook it needs, and the finished
flow as an importable package in two formats:

- `Solution-Lab-NN.zip` — a **Dataverse solution**. Import via *Solutions > Import solution*.
- `Lab-N-….zip` — a **legacy package**. Import via *My flows > Import > Import Package (Legacy)*.

Tenants with *Create in Dataverse solutions* enabled can only use the solution format;
the legacy packages are provided for tenants without that policy.

## Lab environment

The labs run in a dedicated Power Platform environment with Microsoft Dataverse enabled:

```
TGS-2022015539-Applications Integration with Power Apps and Power Automate
```

Check the environment picker in the top-right of the maker portal **before** you build
anything — work created in the wrong environment cannot simply be moved.

Every flow and app is named `Lab N - … (DO NOT DELETE)` so that housekeeping in a shared
training tenant does not remove another learner's work.

## Courseware in this repository

| File | What it is |
|---|---|
| [Slide deck PPTX](courseware/Applications%20Integration%20with%20Power%20Apps%20and%20Power%20Automate-v8.0.pptx) · [PDF](courseware/Applications%20Integration%20with%20Power%20Apps%20and%20Power%20Automate-v8.0.pdf) | Trainer slide deck (108 slides) |
| [Lesson Plan DOCX](courseware/LP-Applications%20Integration%20with%20Power%20Apps%20and%20Power%20Automate.docx) · [PDF](courseware/LP-Applications%20Integration%20with%20Power%20Apps%20and%20Power%20Automate.pdf) | Lesson Plan — 2-day schedule, 480 minutes per day |
| [Learner Guide DOCX](courseware/LG-Applications%20Integration%20with%20Power%20Apps%20and%20Power%20Automate.docx) · [PDF](courseware/LG-Applications%20Integration%20with%20Power%20Apps%20and%20Power%20Automate.pdf) | Learner Guide — the detailed step-by-step for all 14 labs |
| [LG Markdown](LG-Applications%20Integration%20with%20Power%20Apps%20and%20Power%20Automate.md) | Markdown mirror of the Learner Guide |
| `labs/` | 14 lab folders — instructions, data and packages |

The slides carry the concepts and the shape of each lab; the **detailed procedure lives in
the Learner Guide**, not on the slides. The slide deck includes the "Login to LMS and Download
Course Material" slide showing the current lms-tms.tertiaryinfotech.com sign-in page.

## Assessment

| Instrument | Format | Covers |
|---|---|---|
| Written Assessment (SAQ) | 5 open-ended questions · 1 hour · open book | K1–K5 |
| Practical Performance (PP) | 4 scenario tasks · 1.5 hours · open book | A1–A8 |

## Distribution boundary

The slide deck, Learner Guide, Lesson Plan and labs are published in this repository. Assessment
papers, answer keys, the Facilitator Guide, the Assessment Plan, source references and credentials
are confidential and are **not** published; they are distributed through the approved Google Drive /
LMS workflow only.

## Support

Email: enquiry@tertiaryinfotech.com · Tel: +65 6100 0613
Website: [www.tertiarycourses.com.sg](https://www.tertiarycourses.com.sg)
LMS / TMS: [lms-tms.tertiaryinfotech.com](https://lms-tms.tertiaryinfotech.com)

---

Provided by **Tertiary Infotech Academy Pte Ltd** · UEN 201200696W · © 2026 Tertiary Infotech Academy Pte Ltd. All rights reserved.
