# Lab 5 — Conditions and Branching

**Course:** Applications Integration with Power Apps and Power Automate (TGS-2022015539)  
**Topic 2:** Power Automate  
**Maps to:** K2 Proper usage of middleware; A3 Utilise middleware to integrate data and functions across application programs.

---

## What you are building

Real integrations rarely run in a straight line. Add a Condition so urgent service calls are escalated while routine ones are only logged, then add a scheduled digest flow that summarises the day's calls.

**Deliverable —** Flow 'Lab 5 - Conditions and Branching (DO NOT DELETE)' with a working If yes / If no branch, plus a scheduled recurrence flow producing a daily digest.

**Tools —** Power Automate, Office 365 Outlook, Excel Online (Business), Recurrence

> **Environment.** Every lab in this course runs in the Power Platform
> environment **TGS-2022015539-Applications Integration with Power Apps and Power Automate**.
> Check the environment picker in the top-right of the maker portal BEFORE you
> build anything — work created in the wrong environment cannot simply be moved.

> **Name the flow exactly** `Lab 5 - Conditions and Branching (DO NOT DELETE)`.
> The trailing (DO NOT DELETE) marks it as courseware in a shared training tenant.

## Data

- `data/KinetEco Service Calls.xlsx` — upload this to your OneDrive for Business before you start.

Each workbook already contains a real Excel **Table**. Power Apps and Power
Automate can only bind to a named table, never to a loose range — if you build
your own workbook later, remember to Insert > Table first.

## Prebuilt packages

If you want to inspect or restore the finished flow instead of building it:

- `Lab-5-Conditions-and-Branching.zip` — **legacy package**. Import via My flows > Import > Import Package (Legacy).
- `Lab-5b-Daily-Digest.zip` — **legacy package**. Import via My flows > Import > Import Package (Legacy).
- `Solution-Lab-05.zip` — **Dataverse solution**. Import via Solutions > Import solution.

Imported flows arrive **turned off** until you supply your own connection —
open the flow, re-authenticate each connector, then turn it on.

## Steps

1. Save the Lab 4 flow as 'Lab 5 - Conditions and Branching (DO NOT DELETE)'.

2. After the trigger, add a Condition. Set the left side to the Subject dynamic content, the operator to 'contains', and the right side to URGENT.

3. In the If yes branch, add 'Send an email (V2)' to the duty engineer with high importance, and keep the Excel logging step in both branches.

4. In the If no branch, keep only the Excel row so routine calls are recorded without interrupting anyone.

5. Save and test twice: once with URGENT in the subject and once without, confirming a different path each time.

6. Create a second flow, a Scheduled cloud flow named 'Lab 5b - Daily Digest (DO NOT DELETE)', recurring daily at 18:00 Singapore time.

7. In the digest flow add 'List rows present in a table' against ServiceCalls, then a Select and a Create HTML table action.

8. Send the HTML table by email and run the flow manually to verify the digest renders.

## Test it

Run history shows the URGENT test taking the If yes branch and the routine test taking If no. The digest flow sends one email containing an HTML table of the logged calls.

---

◀ Previous: [Lab 4 — Log to Excel](../lab-04/lab-04.md)  
▶ Next: [Lab 6 — Leave Application Approval](../lab-06/lab-06.md)
