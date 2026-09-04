# Lab 2 — Feasibility Scan and Middleware Selection

**Course:** Applications Integration with Power Apps and Power Automate (TGS-2022015539)  
**Topic 1:** Opportunities for Using Power Platform Apps  
**Maps to:** A2 Perform feasibility scan and assessment to identify potential middleware to be used; K1 Types of middleware and their features.

---

## What you are building

Take the top three opportunities from Lab 1 and run a structured feasibility scan against six criteria: connector availability, data readiness, licensing, security and governance, performance/delegation, and maintainability. Decide for each whether Power Platform is the right middleware — and be prepared to justify a 'no'.

**Deliverable —** A Feasibility Assessment Matrix scoring three candidates against six criteria, with a documented Build / Defer / Reject decision and justification for each.

**Tools —** Power Platform admin center, connector catalogue, Microsoft licensing documentation

> **Environment.** Every lab in this course runs in the Power Platform
> environment **TGS-2022015539-Applications Integration with Power Apps and Power Automate**.
> Check the environment picker in the top-right of the maker portal BEFORE you
> build anything — work created in the wrong environment cannot simply be moved.

## Data

- `data/KinetEco Service Calls.xlsx` — upload this to your OneDrive for Business before you start.

Each workbook already contains a real Excel **Table**. Power Apps and Power
Automate can only bind to a named table, never to a loose range — if you build
your own workbook later, remember to Insert > Table first.

## Steps

1. For each of your top three candidates, confirm a standard connector exists — and note whether it is Standard or Premium, because Premium changes the licence needed.
   - <https://make.powerautomate.com/connectors/>

2. Assess data readiness: is the source structured, does it have a stable key, is it an Excel Table (Power Apps cannot bind to a loose range)?

3. Assess security and governance: who owns the data, does the flow need a service account, does a DLP policy separate the connectors you intend to combine?

4. Check the environment's DLP policies in the admin center and record any that would block your design.
   - <https://admin.powerplatform.microsoft.com/>

5. Assess performance: estimate row counts and check each connector's delegation support and throttling limits.

6. Score each criterion 1-5, total the scores, and record a Build / Defer / Reject decision with a one-line justification.

7. Present your top candidate to the class in 2 minutes: the opportunity, the middleware, and why it is feasible.

## Test it

Each of the three candidates has six scores, a total, and a decision with justification. At least one candidate should be Deferred or Rejected with a defensible reason — a scan where everything passes is not a scan.

---

◀ Previous: [Lab 1 — Integration Opportunity Scan](../lab-01/lab-01.md)  
▶ Next: [Lab 3 — Trigger and Actions](../lab-03/lab-03.md)
