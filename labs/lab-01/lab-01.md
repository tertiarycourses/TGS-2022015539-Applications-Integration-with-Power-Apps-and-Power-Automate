# Lab 1 — Integration Opportunity Scan

**Course:** Applications Integration with Power Apps and Power Automate (TGS-2022015539)  
**Topic 1:** Opportunities for Using Power Platform Apps  
**Maps to:** K1 Types of middleware and their features; A1 Identify opportunities for creating connections among various devices, databases, software and applications.

---

## What you are building

Working from the KinetEco service-desk scenario, map the organisation's current manual process end to end and mark every point where data is re-keyed between systems. Each re-keying point is an integration opportunity. Classify each candidate by the systems it joins (device, database, software, application) and score it for volume, error rate and business impact.

**Deliverable —** A completed Integration Opportunity Register listing at least six candidate integrations, each classified and scored.

**Tools —** Power Platform admin center, Excel, Power Automate connector catalogue

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

1. Open the course environment in the maker portal and confirm you are in 'TGS-2022015539-Applications Integration with Power Apps and Power Automate' using the environment picker (top right).

2. Open reference workbook 'KinetEco Service Calls.xlsx' and read the Calls sheet — this is the manual process you are replacing.

3. Draw the current process as a swimlane: Requester -> Service Desk -> Engineer -> Reporting. Mark each hand-off where a human retypes data.

4. For every hand-off, record: source system, target system, data moved, frequency, and what breaks when it is done by hand.

5. Browse the connector catalogue and note which connector would serve each source and target.
   - <https://make.powerautomate.com/connectors/>

6. Classify each opportunity as device / database / software / application integration — this is the K1 middleware-type vocabulary.

7. Score each candidate 1-5 for volume, error rate and business impact; rank the register by total score.

## Test it

Your register lists at least six candidates, each naming a source system, a target system, a connector and a score. The top-ranked candidate should be the service-call intake — the one you automate in Lab 3.

---

▶ Next: [Lab 2 — Feasibility Scan and Middleware Selection](../lab-02/lab-02.md)
