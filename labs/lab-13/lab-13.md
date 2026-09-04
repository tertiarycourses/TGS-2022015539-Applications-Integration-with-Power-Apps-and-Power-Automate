# Lab 13 — End-to-End Approval App

**Course:** Applications Integration with Power Apps and Power Automate (TGS-2022015539)  
**Topic 4:** Integrate Power Apps and Power Automate  
**Maps to:** A3 Utilise middleware to integrate data and functions across application programs; A5 Perform tests and checks on the connections between disparate application programs.

---

## What you are building

Assemble everything into one working solution: the app submits, a flow requests approval, the approver decides in Teams, the outcome is written back, and the app shows live status. This is the integrated system the PP assessment asks you to build.

**Deliverable —** Flow 'Lab 13 - Approval Round Trip (DO NOT DELETE)' plus app 'Lab 13 - Leave Portal App (DO NOT DELETE)' showing live request status from the data source.

**Tools —** Power Apps, Power Automate, Approvals, Microsoft Teams, Excel Online (Business)

> **Environment.** Every lab in this course runs in the Power Platform
> environment **TGS-2022015539-Applications Integration with Power Apps and Power Automate**.
> Check the environment picker in the top-right of the maker portal BEFORE you
> build anything — work created in the wrong environment cannot simply be moved.

> **Name the flow exactly** `Lab 13 - Approval Round Trip (DO NOT DELETE)`.
> The trailing (DO NOT DELETE) marks it as courseware in a shared training tenant.

> **Name the app exactly** `Lab 13 - Leave Portal App (DO NOT DELETE)`.

## Data

- `data/LeaveLog.xlsx` — upload this to your OneDrive for Business before you start.

Each workbook already contains a real Excel **Table**. Power Apps and Power
Automate can only bind to a named table, never to a loose range — if you build
your own workbook later, remember to Insert > Table first.

## Prebuilt packages

If you want to inspect or restore the finished flow instead of building it:

- `Lab-13-Approval-Round-Trip.zip` — **legacy package**. Import via My flows > Import > Import Package (Legacy).
- `Solution-Lab-13.zip` — **Dataverse solution**. Import via Solutions > Import solution.

Imported flows arrive **turned off** until you supply your own connection —
open the flow, re-authenticate each connector, then turn it on.

## Steps

1. Save the Lab 8 app as 'Lab 13 - Leave Portal App (DO NOT DELETE)' and add a third screen scrStatus.

2. Create an Instant cloud flow 'Lab 13 - Approval Round Trip (DO NOT DELETE)' with a Power Apps (V2) trigger taking the five request fields.

3. Add a row to LeaveLog with status Pending, then add 'Start and wait for an approval' assigned to the approver.

4. Add a Condition on the approval Outcome, and in each branch use 'Update a row' to set the LeaveLog status to Approved or Rejected with the approver's comments.

5. Add 'Respond to a Power App or flow' returning the final Status and Comments so the app can confirm immediately.

6. In the app, connect the LeaveLog Excel table as a data source and add a gallery on scrStatus bound to it, filtered to the current user.

   ```
   Filter(LeaveLog, Applicant = txtApplicant.Text)
   ```

7. Add a Refresh button so the learner can pull the latest status after the approver decides.

   ```
   Refresh(LeaveLog)
   ```

8. Publish, submit a request from the app, approve it in Teams, then refresh the status screen.

9. Repeat with a rejection and confirm the status screen and the Excel row both reflect the decision.

## Test it

A request submitted in the app appears as Pending, moves to Approved or Rejected after the approver decides in Teams, and the app's status screen shows the change after Refresh.

---

◀ Previous: [Lab 12 — Return Data from a Flow to the App](../lab-12/lab-12.md)  
▶ Next: [Lab 14 — Troubleshoot and Optimise the Integration](../lab-14/lab-14.md)
