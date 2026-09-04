# Lab 8 — Blank Canvas App with Power Fx

**Course:** Applications Integration with Power Apps and Power Automate (TGS-2022015539)  
**Topic 3:** Power Apps  
**Maps to:** K3 Different types of platforms on which applications run; A6 Verify proper functioning of modules and applications across multiple or integrated platforms.

---

## What you are building

A generated app hides the mechanics. Build one from an empty screen so every control, variable and formula is yours: a leave-request form with validation, variables and navigation, which becomes the front end you wire to a flow in Topic 4.

**Deliverable —** Canvas app 'Lab 8 - Leave Request App (DO NOT DELETE)' — a two-screen app built from blank with validated inputs, variables and navigation.

**Tools —** Power Apps Studio, Power Fx

> **Environment.** Every lab in this course runs in the Power Platform
> environment **TGS-2022015539-Applications Integration with Power Apps and Power Automate**.
> Check the environment picker in the top-right of the maker portal BEFORE you
> build anything — work created in the wrong environment cannot simply be moved.

> **Name the app exactly** `Lab 8 - Leave Request App (DO NOT DELETE)`.

## Data

- `data/Employee Survey.xlsx` — upload this to your OneDrive for Business before you start.

Each workbook already contains a real Excel **Table**. Power Apps and Power
Automate can only bind to a named table, never to a loose range — if you build
your own workbook later, remember to Insert > Table first.

## Steps

1. Create a blank canvas app with Tablet layout named 'Lab 8 - Leave Request App (DO NOT DELETE)'.

2. On Screen1 insert a Label as the title, then Text input controls named txtApplicant and txtReason, a Dropdown ddLeaveType, a Date picker dpStart and a Text input txtDays.

3. Set the dropdown Items to a literal table of leave types.

   ```
   ["Annual","Medical","Childcare","Unpaid"]
   ```

4. Insert a Button named btnSubmit and set its DisplayMode so it is disabled until the form is complete — validation before submission.

   ```
   If(IsBlank(txtApplicant.Text) || IsBlank(txtDays.Text), DisplayMode.Disabled, DisplayMode.Edit)
   ```

5. On btnSubmit OnSelect, set a global variable capturing the request and navigate to a confirmation screen.

   ```
   Set(gblRequest, {Applicant: txtApplicant.Text, LeaveType: ddLeaveType.Selected.Value, Start: dpStart.SelectedDate, Days: Value(txtDays.Text), Reason: txtReason.Text}); Navigate(scrConfirm, ScreenTransition.Cover)
   ```

6. Add a second screen scrConfirm with labels reading back the variable, proving state carried across screens.

   ```
   "Thank you " & gblRequest.Applicant & " - " & gblRequest.Days & " day(s) of " & gblRequest.LeaveType
   ```

7. Add a Back button on the confirmation screen that clears the variable and returns to the form.

   ```
   Set(gblRequest, Blank()); Reset(txtApplicant); Navigate(Screen1)
   ```

8. Use the App checker to clear all formula errors and accessibility warnings.

9. Save and publish the app.

## Test it

The submit button stays disabled until the required fields are filled; submitting navigates to the confirmation screen showing the entered values; App checker reports no errors.

---

◀ Previous: [Lab 7 — Canvas App from Excel Data](../lab-07/lab-07.md)  
▶ Next: [Lab 9 — Custom Connector and API Integration](../lab-09/lab-09.md)
