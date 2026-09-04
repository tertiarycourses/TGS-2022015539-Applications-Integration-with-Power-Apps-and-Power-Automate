# Lab 11 — Call a Flow from a Canvas App

**Course:** Applications Integration with Power Apps and Power Automate (TGS-2022015539)  
**Topic 4:** Integrate Power Apps and Power Automate  
**Maps to:** A4 Support API-level integration; A3 Utilise middleware to integrate data and functions across application programs.

---

## What you are building

Wire the Lab 8 app to a flow. The button no longer just sets a variable — it invokes a Power Automate flow with typed parameters, so the app becomes a front end to an integration rather than a form in isolation.

**Deliverable —** Flow 'Lab 11 - Submit Leave Request (DO NOT DELETE)' triggered from the app, and the Lab 8 app updated to call it with typed parameters.

**Tools —** Power Apps, Power Automate, Power Apps (V2) trigger, Office 365 Outlook

> **Environment.** Every lab in this course runs in the Power Platform
> environment **TGS-2022015539-Applications Integration with Power Apps and Power Automate**.
> Check the environment picker in the top-right of the maker portal BEFORE you
> build anything — work created in the wrong environment cannot simply be moved.

> **Name the flow exactly** `Lab 11 - Submit Leave Request (DO NOT DELETE)`.
> The trailing (DO NOT DELETE) marks it as courseware in a shared training tenant.

> **Name the app exactly** `Lab 8 - Leave Request App (DO NOT DELETE)`.

## Data

- `data/LeaveLog.xlsx` — upload this to your OneDrive for Business before you start.

Each workbook already contains a real Excel **Table**. Power Apps and Power
Automate can only bind to a named table, never to a loose range — if you build
your own workbook later, remember to Insert > Table first.

## Prebuilt packages

If you want to inspect or restore the finished flow instead of building it:

- `Lab-11-Submit-Leave-Request.zip` — **legacy package**. Import via My flows > Import > Import Package (Legacy).
- `Solution-Lab-11.zip` — **Dataverse solution**. Import via Solutions > Import solution.

Imported flows arrive **turned off** until you supply your own connection —
open the flow, re-authenticate each connector, then turn it on.

## Steps

1. In make.powerautomate.com create an Instant cloud flow named 'Lab 11 - Submit Leave Request (DO NOT DELETE)' and choose the trigger 'Power Apps (V2)'.

2. On the trigger add typed inputs matching the app: Applicant (text), LeaveType (text), StartDate (text), Days (number), Reason (text). Typed inputs are what make the call safe.

3. Add 'Send an email (V2)' to the approver, composing the body from the trigger inputs.

4. Add 'Add a row into a table' writing the request into the LeaveLog Excel table with status Submitted.

5. Save the flow, then open the Lab 8 app in Power Apps Studio.

6. On the Power Automate pane choose Add flow and select your Lab 11 flow.

7. Change btnSubmit OnSelect to call the flow with the control values in the trigger's parameter order, then navigate.

   ```
   'Lab11-SubmitLeaveRequest'.Run(txtApplicant.Text, ddLeaveType.Selected.Value, Text(dpStart.SelectedDate), Value(txtDays.Text), txtReason.Text); Navigate(scrConfirm)
   ```

8. Save and publish the app, then submit a real request from the app.

9. Confirm the flow run appears in run history, the email arrives and the Excel row is written.

## Test it

Submitting from the app produces a Succeeded flow run whose trigger inputs exactly match what you typed, plus an approver email and a new LeaveLog row.

---

◀ Previous: [Lab 10 — Cross-Platform Verification and Sharing](../lab-10/lab-10.md)  
▶ Next: [Lab 12 — Return Data from a Flow to the App](../lab-12/lab-12.md)
