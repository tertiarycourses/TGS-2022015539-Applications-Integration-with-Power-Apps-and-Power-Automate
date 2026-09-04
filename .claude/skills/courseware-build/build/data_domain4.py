"""Topic 4 — Integrate Power Apps and Power Automate (K4, A7, A8).

Labs 11-14 join the two halves of the course: the app calls a flow, the flow returns data
to the app, approvals run end to end, and finally the learner deliberately provokes,
diagnoses and fixes the three classes of integration defect.
"""

DOMAIN4 = [
    dict(
        num=11, topic=4,
        title="Call a Flow from a Canvas App",
        objective="A4 Support API-level integration; A3 Utilise middleware to integrate data and functions across application programs.",
        desc=(
            "Wire the Lab 8 app to a flow. The button no longer just sets a variable — it invokes a "
            "Power Automate flow with typed parameters, so the app becomes a front end to an "
            "integration rather than a form in isolation."
        ),
        build="Flow 'Lab 11 - Submit Leave Request (DO NOT DELETE)' triggered from the app, and the Lab 8 app updated to call it with typed parameters.",
        services="Power Apps, Power Automate, Power Apps (V2) trigger, Office 365 Outlook",
        flow_name="Lab 11 - Submit Leave Request (DO NOT DELETE)",
        app_name="Lab 8 - Leave Request App (DO NOT DELETE)",
        environment=True,
        steps=[
            ("In make.powerautomate.com create an Instant cloud flow named 'Lab 11 - Submit Leave Request (DO NOT DELETE)' and choose the trigger 'Power Apps (V2)'.", ""),
            ("On the trigger add typed inputs matching the app: Applicant (text), LeaveType (text), StartDate (text), Days (number), Reason (text). Typed inputs are what make the call safe.", ""),
            ("Add 'Send an email (V2)' to the approver, composing the body from the trigger inputs.", ""),
            ("Add 'Add a row into a table' writing the request into the LeaveLog Excel table with status Submitted.", ""),
            ("Save the flow, then open the Lab 8 app in Power Apps Studio.", ""),
            ("On the Power Automate pane choose Add flow and select your Lab 11 flow.", ""),
            ("Change btnSubmit OnSelect to call the flow with the control values in the trigger's parameter order, then navigate.", "'Lab11-SubmitLeaveRequest'.Run(txtApplicant.Text, ddLeaveType.Selected.Value, Text(dpStart.SelectedDate), Value(txtDays.Text), txtReason.Text); Navigate(scrConfirm)"),
            ("Save and publish the app, then submit a real request from the app.", ""),
            ("Confirm the flow run appears in run history, the email arrives and the Excel row is written.", ""),
        ],
        test="Submitting from the app produces a Succeeded flow run whose trigger inputs exactly match what you typed, plus an approver email and a new LeaveLog row.",
        stages=[('Create the instant flow', 'Power Apps (V2) trigger'),
                ('Declare typed inputs', 'text and number'),
                ('Add email and Excel', 'approver plus log row'),
                ('Add the flow to the app', 'Power Automate pane'),
                ('Call it from the button', 'pass control values')],
    ),
    dict(
        num=12, topic=4,
        title="Return Data from a Flow to the App",
        objective="A4 Support API-level integration; A6 Verify proper functioning of modules and applications across multiple or integrated platforms.",
        desc=(
            "So far data has flowed one way. Close the round trip with 'Respond to a Power App or flow' "
            "so the flow returns a result the app displays — the pattern that lets an app use server-side "
            "logic, secured credentials or external APIs it cannot reach directly."
        ),
        build="Flow 'Lab 12 - Return Leave Balance (DO NOT DELETE)' returning typed values, and the app displaying the returned balance and reference number.",
        services="Power Apps, Power Automate, Respond to a Power App or flow, Excel Online (Business)",
        flow_name="Lab 12 - Return Leave Balance (DO NOT DELETE)",
        environment=True,
        steps=[
            ("Create an Instant cloud flow 'Lab 12 - Return Leave Balance (DO NOT DELETE)' with a Power Apps (V2) trigger taking Applicant (text) and Days (number).", ""),
            ("Add 'List rows present in a table' against the LeaveLog table, filtering to the applicant.", ""),
            ("Add a Compose that computes the remaining balance from an annual entitlement of 14 days.", "sub(14, int(triggerBody()['number']))"),
            ("Add 'Respond to a Power App or flow' with outputs Balance (number), Reference (text) and Status (text).", ""),
            ("Set Reference to a generated unique value so the app receives something a user can quote.", "concat('LV-', utcNow('yyyyMMdd'), '-', rand(1000,9999))"),
            ("Save the flow, return to the app, and add this flow alongside the Lab 11 flow.", ""),
            ("Change the submit button to capture the flow's return value into a variable instead of discarding it.", "Set(gblResult, 'Lab12-ReturnLeaveBalance'.Run(txtApplicant.Text, Value(txtDays.Text))); Navigate(scrConfirm)"),
            ("On the confirmation screen add labels bound to the returned fields.", "\"Reference: \" & gblResult.reference & \"  |  Balance: \" & gblResult.balance & \" days\""),
            ("Publish and submit a request, confirming the reference and balance shown in the app match the flow's run outputs.", ""),
        ],
        test="The confirmation screen displays a reference number and remaining balance produced by the flow, and those values match the 'Respond to a Power App or flow' outputs in run history.",
        stages=[('Create the return flow', 'Power Apps (V2) trigger'),
                ('Compute the balance', 'expression on inputs'),
                ('Generate a reference', 'concat and rand'),
                ('Respond to the app', 'typed outputs'),
                ('Show it in the app', 'bind to the result')],
    ),
    dict(
        num=13, topic=4,
        title="End-to-End Approval App",
        objective="A3 Utilise middleware to integrate data and functions across application programs; A5 Perform tests and checks on the connections between disparate application programs.",
        desc=(
            "Assemble everything into one working solution: the app submits, a flow requests approval, "
            "the approver decides in Teams, the outcome is written back, and the app shows live status. "
            "This is the integrated system the PP assessment asks you to build."
        ),
        build="Flow 'Lab 13 - Approval Round Trip (DO NOT DELETE)' plus app 'Lab 13 - Leave Portal App (DO NOT DELETE)' showing live request status from the data source.",
        services="Power Apps, Power Automate, Approvals, Microsoft Teams, Excel Online (Business)",
        flow_name="Lab 13 - Approval Round Trip (DO NOT DELETE)",
        app_name="Lab 13 - Leave Portal App (DO NOT DELETE)",
        environment=True,
        steps=[
            ("Save the Lab 8 app as 'Lab 13 - Leave Portal App (DO NOT DELETE)' and add a third screen scrStatus.", ""),
            ("Create an Instant cloud flow 'Lab 13 - Approval Round Trip (DO NOT DELETE)' with a Power Apps (V2) trigger taking the five request fields.", ""),
            ("Add a row to LeaveLog with status Pending, then add 'Start and wait for an approval' assigned to the approver.", ""),
            ("Add a Condition on the approval Outcome, and in each branch use 'Update a row' to set the LeaveLog status to Approved or Rejected with the approver's comments.", ""),
            ("Add 'Respond to a Power App or flow' returning the final Status and Comments so the app can confirm immediately.", ""),
            ("In the app, connect the LeaveLog Excel table as a data source and add a gallery on scrStatus bound to it, filtered to the current user.", "Filter(LeaveLog, Applicant = txtApplicant.Text)"),
            ("Add a Refresh button so the learner can pull the latest status after the approver decides.", "Refresh(LeaveLog)"),
            ("Publish, submit a request from the app, approve it in Teams, then refresh the status screen.", ""),
            ("Repeat with a rejection and confirm the status screen and the Excel row both reflect the decision.", ""),
        ],
        test="A request submitted in the app appears as Pending, moves to Approved or Rejected after the approver decides in Teams, and the app's status screen shows the change after Refresh.",
        stages=[('Copy the app, add a screen', 'status screen'),
                ('Create the round-trip flow', 'five request fields'),
                ('Log as Pending', 'then wait for approval'),
                ('Update on the outcome', 'Approved or Rejected'),
                ('Refresh and verify', 'status shown in the app')],
    ),
    dict(
        num=14, topic=4,
        title="Troubleshoot and Optimise the Integration",
        objective="K4 Potential technical, compatibility or performance issues in application integration; A7 Highlight technical, compatibility or performance issues; A8 Implement modifications to mitigate the issues identified.",
        desc=(
            "Deliberately break your working solution in three ways — technical, compatibility and "
            "performance — then diagnose each from the evidence, fix it, and record the finding. This is "
            "the A7/A8 discipline the assessment tests, and it is far easier to learn on a system you built."
        ),
        build="A completed Integration Issue Log documenting three provoked defects with evidence, root cause and the implemented fix, plus the repaired solution.",
        services="Power Automate run history, Power Apps App checker, Monitor, delegation warnings",
        environment=True,
        steps=[
            ("TECHNICAL — open the Lab 13 flow's connections and delete or invalidate the Excel connection, then run the app. Capture the exact error from run history.", ""),
            ("Diagnose it as an authentication/connection failure, repair the connection, and re-run to prove the fix.", ""),
            ("COMPATIBILITY — in the app pass Days as text rather than a number to the flow, and observe the schema-validation failure.", "'Lab13-ApprovalRoundTrip'.Run(txtApplicant.Text, ddLeaveType.Selected.Value, Text(dpStart.SelectedDate), txtDays.Text, txtReason.Text)"),
            ("Fix it by coercing the type at the call site, and note the general rule that connector inputs are typed contracts.", "Value(txtDays.Text)"),
            ("PERFORMANCE — add a gallery over a large data source using a non-delegable function and read the blue delegation warning.", "Filter(LeaveLog, Len(Applicant) > 3)"),
            ("Explain why only the first 500 (or 2000) rows are processed, then rewrite it to a delegable filter and confirm the warning clears.", "Filter(LeaveLog, StartsWith(Applicant, txtSearch.Text))"),
            ("Raise the data row limit in Settings > General and explain why that is a mitigation but not a cure.", ""),
            ("Record all three in an Issue Log: symptom, evidence, classification, root cause, fix implemented, and how you verified the fix.", ""),
            ("Re-run the full end-to-end scenario from Lab 13 to prove the repaired solution still works.", ""),
        ],
        test="The Issue Log documents three issues, one of each class, each with real evidence, a root cause and a verified fix. The end-to-end approval scenario runs successfully after all repairs.",
        stages=[('Break the connection', 'technical fault'),
                ('Diagnose from run history', 'read the real error'),
                ('Pass the wrong type', 'compatibility fault'),
                ('Delegation warning', 'performance fault'),
                ('Fix, verify and log', 'three issues documented')],
    ),
]
