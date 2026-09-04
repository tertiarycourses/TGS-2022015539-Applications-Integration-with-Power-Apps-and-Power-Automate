"""Topic 2 — Power Automate (K2, A3, A5).

Labs 3-6 build automation capability in strict order: a first trigger/action pair,
then data landing in a real store, then branching logic, then human approval.
Each lab reuses the artefact built in the previous one.
"""

DOMAIN2 = [
    dict(
        num=3, topic=2,
        title="Trigger and Actions",
        objective="K2 Proper usage of middleware; A3 Utilise middleware to integrate data and functions across application programs.",
        desc=(
            "Build your first cloud flow: an automated flow that fires when a new service-call item "
            "arrives and sends a formatted notification. This is the atom of every integration — one "
            "trigger, one or more actions, and dynamic content carrying data between them."
        ),
        build="Flow 'Lab 3 - Trigger and Actions (DO NOT DELETE)' — an automated cloud flow with a trigger, a compose step and a notification action, tested with a successful run.",
        services="Power Automate, Office 365 Outlook, Compose",
        flow_name="Lab 3 - Trigger and Actions (DO NOT DELETE)",
        environment=True,
        steps=[
            ("In make.powerautomate.com confirm the environment picker reads the course environment, then choose Create > Automated cloud flow.", "https://make.powerautomate.com"),
            ("Name the flow exactly 'Lab 3 - Trigger and Actions (DO NOT DELETE)' and pick the trigger 'When a new email arrives (V3)'. Click Create.", ""),
            ("On the trigger, set Folder to Inbox and expand Advanced options: set Subject Filter to SERVICE to narrow what fires the flow.", ""),
            ("Add a Compose action. In its Inputs, insert dynamic content From, Subject and Received Time — this proves data crosses from the trigger into the next action.", ""),
            ("Add 'Send an email notification (V3)'. Set To to your own address, Subject to 'Service call logged', and put the Compose Outputs in the Body.", ""),
            ("Click Flow checker (top right) and clear every error and warning before saving.", ""),
            ("Save, then click Test > Manually, and send yourself an email whose subject contains SERVICE.", ""),
            ("Open the run in the 28-day run history and expand each step to read its raw inputs and outputs — this is how you evidence a working connection.", ""),
        ],
        test="The run history shows one Succeeded run. Expanding the Compose step shows the sender, subject and timestamp carried from the trigger, and the notification email arrives in your inbox.",
        stages=[('Create an automated flow', 'new cloud flow'),
                ('Pick the email trigger', 'When a new email arrives'),
                ('Filter what fires it', 'subject contains SERVICE'),
                ('Compose the details', 'from · subject · time'),
                ('Send the notification', 'and test the run')],
    ),
    dict(
        num=4, topic=2,
        title="Log to Excel",
        objective="A3 Utilise middleware to integrate data and functions across application programs; A5 Perform tests and checks on the connections between disparate application programs.",
        desc=(
            "Extend Lab 3 so the service call is not merely announced but recorded. The flow writes a "
            "row into an Excel Online table in OneDrive, turning a notification into a durable "
            "integration between a mail system and a data store."
        ),
        build="Flow 'Lab 4 - Log to Excel (DO NOT DELETE)' writing one row per service call into a formatted Excel Table, verified by inspecting the workbook.",
        services="Power Automate, Office 365 Outlook, Excel Online (Business), OneDrive for Business",
        flow_name="Lab 4 - Log to Excel (DO NOT DELETE)",
        environment=True,
        steps=[
            ("Upload 'KinetEco Service Calls.xlsx' from the lab data folder to your OneDrive for Business.", ""),
            ("Open the workbook, select the Calls range and choose Insert > Table with 'My table has headers' ticked. Power Apps and Power Automate can only bind to a TABLE, never a loose range — this is the single most common beginner failure.", ""),
            ("Rename the table ServiceCalls in Table Design, then save and close the workbook.", ""),
            ("Copy the Lab 3 flow using Save As, and rename the copy exactly 'Lab 4 - Log to Excel (DO NOT DELETE)'.", ""),
            ("Add the action 'Add a row into a table' (Excel Online Business). Pick the OneDrive location, the workbook and the ServiceCalls table.", ""),
            ("Map each column to dynamic content: Reported By to From, Problem to Subject, Date Reported to Received Time.", ""),
            ("Run Flow checker, save, and test the flow by sending another SERVICE email.", ""),
            ("Open the workbook and confirm the new row. Then break the integration on purpose: rename the table in Excel, re-run, and read the error the flow returns.", ""),
            ("Restore the table name and re-run to confirm the flow recovers — you have now both tested and diagnosed a connection.", ""),
        ],
        test="A new row appears in the ServiceCalls table for each test email. The deliberate rename produces a clear failure in run history, and restoring the name returns the flow to Succeeded.",
        stages=[('Format data as a Table', 'Insert > Table'),
                ('Name the table', 'ServiceCalls'),
                ('Copy the Lab 3 flow', 'Save As'),
                ('Add a row to the table', 'map dynamic content'),
                ('Break it and fix it', 'rename, re-run, restore')],
    ),
    dict(
        num=5, topic=2,
        title="Conditions and Branching",
        objective="K2 Proper usage of middleware; A3 Utilise middleware to integrate data and functions across application programs.",
        desc=(
            "Real integrations rarely run in a straight line. Add a Condition so urgent service calls are "
            "escalated while routine ones are only logged, then add a scheduled digest flow that "
            "summarises the day's calls."
        ),
        build="Flow 'Lab 5 - Conditions and Branching (DO NOT DELETE)' with a working If yes / If no branch, plus a scheduled recurrence flow producing a daily digest.",
        services="Power Automate, Office 365 Outlook, Excel Online (Business), Recurrence",
        flow_name="Lab 5 - Conditions and Branching (DO NOT DELETE)",
        environment=True,
        steps=[
            ("Save the Lab 4 flow as 'Lab 5 - Conditions and Branching (DO NOT DELETE)'.", ""),
            ("After the trigger, add a Condition. Set the left side to the Subject dynamic content, the operator to 'contains', and the right side to URGENT.", ""),
            ("In the If yes branch, add 'Send an email (V2)' to the duty engineer with high importance, and keep the Excel logging step in both branches.", ""),
            ("In the If no branch, keep only the Excel row so routine calls are recorded without interrupting anyone.", ""),
            ("Save and test twice: once with URGENT in the subject and once without, confirming a different path each time.", ""),
            ("Create a second flow, a Scheduled cloud flow named 'Lab 5b - Daily Digest (DO NOT DELETE)', recurring daily at 18:00 Singapore time.", ""),
            ("In the digest flow add 'List rows present in a table' against ServiceCalls, then a Select and a Create HTML table action.", ""),
            ("Send the HTML table by email and run the flow manually to verify the digest renders.", ""),
        ],
        test="Run history shows the URGENT test taking the If yes branch and the routine test taking If no. The digest flow sends one email containing an HTML table of the logged calls.",
        stages=[('Copy the Lab 4 flow', 'Save As'),
                ('Add a Condition', 'subject contains URGENT'),
                ('Escalate on the yes branch', 'high-importance email'),
                ('Log only on the no branch', 'routine calls'),
                ('Add a scheduled digest', 'daily at 18:00')],
    ),
    dict(
        num=6, topic=2,
        title="Leave Application Approval",
        objective="A3 Utilise middleware to integrate data and functions across application programs; A5 Perform tests and checks on the connections between disparate application programs.",
        desc=(
            "Add the human to the loop. Build an approval flow where a leave request is routed to an "
            "approver, the decision branches the flow, and the outcome is written back and notified — "
            "the canonical human-in-the-loop integration pattern."
        ),
        build="Flow 'Lab 6 - Leave Application Approval (DO NOT DELETE)' with a Start and wait for an approval action, an outcome condition, and write-back to Excel.",
        services="Power Automate, Approvals, Office 365 Outlook, Excel Online (Business), Microsoft Teams",
        flow_name="Lab 6 - Leave Application Approval (DO NOT DELETE)",
        environment=True,
        steps=[
            ("Create a new Instant cloud flow named 'Lab 6 - Leave Application Approval (DO NOT DELETE)' triggered manually, so you can test it repeatedly without waiting for an event.", ""),
            ("Add inputs to the trigger: Applicant (text), Leave Type (text), Start Date (date), Days (number), Reason (text).", ""),
            ("Add 'Start and wait for an approval'. Choose 'Approve/Reject - First to respond', and compose the Title and Details from the trigger inputs.", ""),
            ("Set Assigned to to your own address so you can approve your own test runs. Approvals arrive in Teams and the Approvals app.", ""),
            ("Add a Condition testing whether the Outcome dynamic content is equal to Approve.", ""),
            ("In If yes, send a confirmation email to the applicant and add a row to an Excel LeaveLog table with status Approved.", ""),
            ("In If no, send a rejection email including the approver's Comments, and log the row with status Rejected.", ""),
            ("Save and run the flow, approve the request when it arrives, then run again and reject it.", ""),
            ("Inspect both runs in run history and confirm the branch, the email and the Excel row match the decision each time.", ""),
        ],
        test="Two runs complete: one Approved and one Rejected. Each produced the correct branch, the correct email, and a correctly-statused row in the LeaveLog table.",
        stages=[('Create an instant flow', 'manual trigger'),
                ('Add the request inputs', 'applicant · type · days'),
                ('Start and wait for approval', 'first to respond'),
                ('Branch on the outcome', 'Approve or Reject'),
                ('Write back and notify', 'Excel row plus email')],
    ),
]
