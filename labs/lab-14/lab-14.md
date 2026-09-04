# Lab 14 — Troubleshoot and Optimise the Integration

**Course:** Applications Integration with Power Apps and Power Automate (TGS-2022015539)  
**Topic 4:** Integrate Power Apps and Power Automate  
**Maps to:** K4 Potential technical, compatibility or performance issues in application integration; A7 Highlight technical, compatibility or performance issues; A8 Implement modifications to mitigate the issues identified.

---

## What you are building

Deliberately break your working solution in three ways — technical, compatibility and performance — then diagnose each from the evidence, fix it, and record the finding. This is the A7/A8 discipline the assessment tests, and it is far easier to learn on a system you built.

**Deliverable —** A completed Integration Issue Log documenting three provoked defects with evidence, root cause and the implemented fix, plus the repaired solution.

**Tools —** Power Automate run history, Power Apps App checker, Monitor, delegation warnings

> **Environment.** Every lab in this course runs in the Power Platform
> environment **TGS-2022015539-Applications Integration with Power Apps and Power Automate**.
> Check the environment picker in the top-right of the maker portal BEFORE you
> build anything — work created in the wrong environment cannot simply be moved.

## Data

- `data/LeaveLog.xlsx` — upload this to your OneDrive for Business before you start.

Each workbook already contains a real Excel **Table**. Power Apps and Power
Automate can only bind to a named table, never to a loose range — if you build
your own workbook later, remember to Insert > Table first.

## Steps

1. TECHNICAL — open the Lab 13 flow's connections and delete or invalidate the Excel connection, then run the app. Capture the exact error from run history.

2. Diagnose it as an authentication/connection failure, repair the connection, and re-run to prove the fix.

3. COMPATIBILITY — in the app pass Days as text rather than a number to the flow, and observe the schema-validation failure.

   ```
   'Lab13-ApprovalRoundTrip'.Run(txtApplicant.Text, ddLeaveType.Selected.Value, Text(dpStart.SelectedDate), txtDays.Text, txtReason.Text)
   ```

4. Fix it by coercing the type at the call site, and note the general rule that connector inputs are typed contracts.

   ```
   Value(txtDays.Text)
   ```

5. PERFORMANCE — add a gallery over a large data source using a non-delegable function and read the blue delegation warning.

   ```
   Filter(LeaveLog, Len(Applicant) > 3)
   ```

6. Explain why only the first 500 (or 2000) rows are processed, then rewrite it to a delegable filter and confirm the warning clears.

   ```
   Filter(LeaveLog, StartsWith(Applicant, txtSearch.Text))
   ```

7. Raise the data row limit in Settings > General and explain why that is a mitigation but not a cure.

8. Record all three in an Issue Log: symptom, evidence, classification, root cause, fix implemented, and how you verified the fix.

9. Re-run the full end-to-end scenario from Lab 13 to prove the repaired solution still works.

## Test it

The Issue Log documents three issues, one of each class, each with real evidence, a root cause and a verified fix. The end-to-end approval scenario runs successfully after all repairs.

---

◀ Previous: [Lab 13 — End-to-End Approval App](../lab-13/lab-13.md)  
