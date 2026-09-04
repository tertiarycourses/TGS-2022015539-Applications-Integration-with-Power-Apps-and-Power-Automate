# Lab 4 — Log to Excel

**Course:** Applications Integration with Power Apps and Power Automate (TGS-2022015539)  
**Topic 2:** Power Automate  
**Maps to:** A3 Utilise middleware to integrate data and functions across application programs; A5 Perform tests and checks on the connections between disparate application programs.

---

## What you are building

Extend Lab 3 so the service call is not merely announced but recorded. The flow writes a row into an Excel Online table in OneDrive, turning a notification into a durable integration between a mail system and a data store.

**Deliverable —** Flow 'Lab 4 - Log to Excel (DO NOT DELETE)' writing one row per service call into a formatted Excel Table, verified by inspecting the workbook.

**Tools —** Power Automate, Office 365 Outlook, Excel Online (Business), OneDrive for Business

> **Environment.** Every lab in this course runs in the Power Platform
> environment **TGS-2022015539-Applications Integration with Power Apps and Power Automate**.
> Check the environment picker in the top-right of the maker portal BEFORE you
> build anything — work created in the wrong environment cannot simply be moved.

> **Name the flow exactly** `Lab 4 - Log to Excel (DO NOT DELETE)`.
> The trailing (DO NOT DELETE) marks it as courseware in a shared training tenant.

## Data

- `data/KinetEco Service Calls.xlsx` — upload this to your OneDrive for Business before you start.

Each workbook already contains a real Excel **Table**. Power Apps and Power
Automate can only bind to a named table, never to a loose range — if you build
your own workbook later, remember to Insert > Table first.

## Prebuilt packages

If you want to inspect or restore the finished flow instead of building it:

- `Lab-4-Log-to-Excel.zip` — **legacy package**. Import via My flows > Import > Import Package (Legacy).
- `Solution-Lab-04.zip` — **Dataverse solution**. Import via Solutions > Import solution.

Imported flows arrive **turned off** until you supply your own connection —
open the flow, re-authenticate each connector, then turn it on.

## Steps

1. Upload 'KinetEco Service Calls.xlsx' from the lab data folder to your OneDrive for Business.

2. Open the workbook, select the Calls range and choose Insert > Table with 'My table has headers' ticked. Power Apps and Power Automate can only bind to a TABLE, never a loose range — this is the single most common beginner failure.

3. Rename the table ServiceCalls in Table Design, then save and close the workbook.

4. Copy the Lab 3 flow using Save As, and rename the copy exactly 'Lab 4 - Log to Excel (DO NOT DELETE)'.

5. Add the action 'Add a row into a table' (Excel Online Business). Pick the OneDrive location, the workbook and the ServiceCalls table.

6. Map each column to dynamic content: Reported By to From, Problem to Subject, Date Reported to Received Time.

7. Run Flow checker, save, and test the flow by sending another SERVICE email.

8. Open the workbook and confirm the new row. Then break the integration on purpose: rename the table in Excel, re-run, and read the error the flow returns.

9. Restore the table name and re-run to confirm the flow recovers — you have now both tested and diagnosed a connection.

## Test it

A new row appears in the ServiceCalls table for each test email. The deliberate rename produces a clear failure in run history, and restoring the name returns the flow to Succeeded.

---

◀ Previous: [Lab 3 — Trigger and Actions](../lab-03/lab-03.md)  
▶ Next: [Lab 5 — Conditions and Branching](../lab-05/lab-05.md)
