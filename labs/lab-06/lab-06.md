# Lab 6 — Leave Application Approval

**Course:** Applications Integration with Power Apps and Power Automate (TGS-2022015539)  
**Topic 2:** Power Automate  
**Maps to:** A3 Utilise middleware to integrate data and functions across application programs; A5 Perform tests and checks on the connections between disparate application programs.

---

## What you are building

Add the human to the loop. Build an approval flow where a leave request is routed to an approver, the decision branches the flow, and the outcome is written back and notified — the canonical human-in-the-loop integration pattern.

**Deliverable —** Flow 'Lab 6 - Leave Application Approval (DO NOT DELETE)' with a Start and wait for an approval action, an outcome condition, and write-back to Excel.

**Tools —** Power Automate, Approvals, Office 365 Outlook, Excel Online (Business), Microsoft Teams

> **Environment.** Every lab in this course runs in the Power Platform
> environment **TGS-2022015539-Applications Integration with Power Apps and Power Automate**.
> Check the environment picker in the top-right of the maker portal BEFORE you
> build anything — work created in the wrong environment cannot simply be moved.

> **Name the flow exactly** `Lab 6 - Leave Application Approval (DO NOT DELETE)`.
> The trailing (DO NOT DELETE) marks it as courseware in a shared training tenant.

## Data

- `data/LeaveLog.xlsx` — upload this to your OneDrive for Business before you start.

Each workbook already contains a real Excel **Table**. Power Apps and Power
Automate can only bind to a named table, never to a loose range — if you build
your own workbook later, remember to Insert > Table first.

## Prebuilt packages

If you want to inspect or restore the finished flow instead of building it:

- `Lab-6-Leave-Application-Approval.zip` — **legacy package**. Import via My flows > Import > Import Package (Legacy).
- `Solution-Lab-06.zip` — **Dataverse solution**. Import via Solutions > Import solution.

Imported flows arrive **turned off** until you supply your own connection —
open the flow, re-authenticate each connector, then turn it on.

## Steps

1. Create a new Instant cloud flow named 'Lab 6 - Leave Application Approval (DO NOT DELETE)' triggered manually, so you can test it repeatedly without waiting for an event.

2. Add inputs to the trigger: Applicant (text), Leave Type (text), Start Date (date), Days (number), Reason (text).

3. Add 'Start and wait for an approval'. Choose 'Approve/Reject - First to respond', and compose the Title and Details from the trigger inputs.

4. Set Assigned to to your own address so you can approve your own test runs. Approvals arrive in Teams and the Approvals app.

5. Add a Condition testing whether the Outcome dynamic content is equal to Approve.

6. In If yes, send a confirmation email to the applicant and add a row to an Excel LeaveLog table with status Approved.

7. In If no, send a rejection email including the approver's Comments, and log the row with status Rejected.

8. Save and run the flow, approve the request when it arrives, then run again and reject it.

9. Inspect both runs in run history and confirm the branch, the email and the Excel row match the decision each time.

## Test it

Two runs complete: one Approved and one Rejected. Each produced the correct branch, the correct email, and a correctly-statused row in the LeaveLog table.

---

◀ Previous: [Lab 5 — Conditions and Branching](../lab-05/lab-05.md)  
▶ Next: [Lab 7 — Canvas App from Excel Data](../lab-07/lab-07.md)
