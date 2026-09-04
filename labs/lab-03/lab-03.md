# Lab 3 — Trigger and Actions

**Course:** Applications Integration with Power Apps and Power Automate (TGS-2022015539)  
**Topic 2:** Power Automate  
**Maps to:** K2 Proper usage of middleware; A3 Utilise middleware to integrate data and functions across application programs.

---

## What you are building

Build your first cloud flow: an automated flow that fires when a new service-call item arrives and sends a formatted notification. This is the atom of every integration — one trigger, one or more actions, and dynamic content carrying data between them.

**Deliverable —** Flow 'Lab 3 - Trigger and Actions (DO NOT DELETE)' — an automated cloud flow with a trigger, a compose step and a notification action, tested with a successful run.

**Tools —** Power Automate, Office 365 Outlook, Compose

> **Environment.** Every lab in this course runs in the Power Platform
> environment **TGS-2022015539-Applications Integration with Power Apps and Power Automate**.
> Check the environment picker in the top-right of the maker portal BEFORE you
> build anything — work created in the wrong environment cannot simply be moved.

> **Name the flow exactly** `Lab 3 - Trigger and Actions (DO NOT DELETE)`.
> The trailing (DO NOT DELETE) marks it as courseware in a shared training tenant.

## Data

- `data/KinetEco Service Calls.xlsx` — upload this to your OneDrive for Business before you start.

Each workbook already contains a real Excel **Table**. Power Apps and Power
Automate can only bind to a named table, never to a loose range — if you build
your own workbook later, remember to Insert > Table first.

## Prebuilt packages

If you want to inspect or restore the finished flow instead of building it:

- `Lab-3-Trigger-and-Actions.zip` — **legacy package**. Import via My flows > Import > Import Package (Legacy).
- `Solution-Lab-03.zip` — **Dataverse solution**. Import via Solutions > Import solution.

Imported flows arrive **turned off** until you supply your own connection —
open the flow, re-authenticate each connector, then turn it on.

## Steps

1. In make.powerautomate.com confirm the environment picker reads the course environment, then choose Create > Automated cloud flow.
   - <https://make.powerautomate.com>

2. Name the flow exactly 'Lab 3 - Trigger and Actions (DO NOT DELETE)' and pick the trigger 'When a new email arrives (V3)'. Click Create.

3. On the trigger, set Folder to Inbox and expand Advanced options: set Subject Filter to SERVICE to narrow what fires the flow.

4. Add a Compose action. In its Inputs, insert dynamic content From, Subject and Received Time — this proves data crosses from the trigger into the next action.

5. Add 'Send an email notification (V3)'. Set To to your own address, Subject to 'Service call logged', and put the Compose Outputs in the Body.

6. Click Flow checker (top right) and clear every error and warning before saving.

7. Save, then click Test > Manually, and send yourself an email whose subject contains SERVICE.

8. Open the run in the 28-day run history and expand each step to read its raw inputs and outputs — this is how you evidence a working connection.

## Test it

The run history shows one Succeeded run. Expanding the Compose step shows the sender, subject and timestamp carried from the trigger, and the notification email arrives in your inbox.

---

◀ Previous: [Lab 2 — Feasibility Scan and Middleware Selection](../lab-02/lab-02.md)  
▶ Next: [Lab 4 — Log to Excel](../lab-04/lab-04.md)
