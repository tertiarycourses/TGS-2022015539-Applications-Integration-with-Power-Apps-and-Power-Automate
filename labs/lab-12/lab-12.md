# Lab 12 — Return Data from a Flow to the App

**Course:** Applications Integration with Power Apps and Power Automate (TGS-2022015539)  
**Topic 4:** Integrate Power Apps and Power Automate  
**Maps to:** A4 Support API-level integration; A6 Verify proper functioning of modules and applications across multiple or integrated platforms.

---

## What you are building

So far data has flowed one way. Close the round trip with 'Respond to a Power App or flow' so the flow returns a result the app displays — the pattern that lets an app use server-side logic, secured credentials or external APIs it cannot reach directly.

**Deliverable —** Flow 'Lab 12 - Return Leave Balance (DO NOT DELETE)' returning typed values, and the app displaying the returned balance and reference number.

**Tools —** Power Apps, Power Automate, Respond to a Power App or flow, Excel Online (Business)

> **Environment.** Every lab in this course runs in the Power Platform
> environment **TGS-2022015539-Applications Integration with Power Apps and Power Automate**.
> Check the environment picker in the top-right of the maker portal BEFORE you
> build anything — work created in the wrong environment cannot simply be moved.

> **Name the flow exactly** `Lab 12 - Return Leave Balance (DO NOT DELETE)`.
> The trailing (DO NOT DELETE) marks it as courseware in a shared training tenant.

## Data

- `data/LeaveLog.xlsx` — upload this to your OneDrive for Business before you start.

Each workbook already contains a real Excel **Table**. Power Apps and Power
Automate can only bind to a named table, never to a loose range — if you build
your own workbook later, remember to Insert > Table first.

## Prebuilt packages

If you want to inspect or restore the finished flow instead of building it:

- `Lab-12-Return-Leave-Balance.zip` — **legacy package**. Import via My flows > Import > Import Package (Legacy).
- `Solution-Lab-12.zip` — **Dataverse solution**. Import via Solutions > Import solution.

Imported flows arrive **turned off** until you supply your own connection —
open the flow, re-authenticate each connector, then turn it on.

## Steps

1. Create an Instant cloud flow 'Lab 12 - Return Leave Balance (DO NOT DELETE)' with a Power Apps (V2) trigger taking Applicant (text) and Days (number).

2. Add 'List rows present in a table' against the LeaveLog table, filtering to the applicant.

3. Add a Compose that computes the remaining balance from an annual entitlement of 14 days.

   ```
   sub(14, int(triggerBody()['number']))
   ```

4. Add 'Respond to a Power App or flow' with outputs Balance (number), Reference (text) and Status (text).

5. Set Reference to a generated unique value so the app receives something a user can quote.

   ```
   concat('LV-', utcNow('yyyyMMdd'), '-', rand(1000,9999))
   ```

6. Save the flow, return to the app, and add this flow alongside the Lab 11 flow.

7. Change the submit button to capture the flow's return value into a variable instead of discarding it.

   ```
   Set(gblResult, 'Lab12-ReturnLeaveBalance'.Run(txtApplicant.Text, Value(txtDays.Text))); Navigate(scrConfirm)
   ```

8. On the confirmation screen add labels bound to the returned fields.

   ```
   "Reference: " & gblResult.reference & "  |  Balance: " & gblResult.balance & " days"
   ```

9. Publish and submit a request, confirming the reference and balance shown in the app match the flow's run outputs.

## Test it

The confirmation screen displays a reference number and remaining balance produced by the flow, and those values match the 'Respond to a Power App or flow' outputs in run history.

---

◀ Previous: [Lab 11 — Call a Flow from a Canvas App](../lab-11/lab-11.md)  
▶ Next: [Lab 13 — End-to-End Approval App](../lab-13/lab-13.md)
