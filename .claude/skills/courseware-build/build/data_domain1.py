"""Topic 1 — Opportunities for Using Power Platform Apps (K1, A1, A2).

Labs 1-2 establish the integration landscape and the feasibility discipline BEFORE
any building happens, so every later lab is a deliberate choice rather than a demo.
"""

DOMAIN1 = [
    dict(
        num=1, topic=1,
        title="Integration Opportunity Scan",
        objective="K1 Types of middleware and their features; A1 Identify opportunities for creating connections among various devices, databases, software and applications.",
        desc=(
            "Working from the KinetEco service-desk scenario, map the organisation's current manual "
            "process end to end and mark every point where data is re-keyed between systems. Each "
            "re-keying point is an integration opportunity. Classify each candidate by the systems it "
            "joins (device, database, software, application) and score it for volume, error rate and "
            "business impact."
        ),
        build="A completed Integration Opportunity Register listing at least six candidate integrations, each classified and scored.",
        services="Power Platform admin center, Excel, Power Automate connector catalogue",
        environment=True,
        steps=[
            ("Open the course environment in the maker portal and confirm you are in 'TGS-2022015539-Applications Integration with Power Apps and Power Automate' using the environment picker (top right).", ""),
            ("Open reference workbook 'KinetEco Service Calls.xlsx' and read the Calls sheet — this is the manual process you are replacing.", ""),
            ("Draw the current process as a swimlane: Requester -> Service Desk -> Engineer -> Reporting. Mark each hand-off where a human retypes data.", ""),
            ("For every hand-off, record: source system, target system, data moved, frequency, and what breaks when it is done by hand.", ""),
            ("Browse the connector catalogue and note which connector would serve each source and target.", "https://make.powerautomate.com/connectors/"),
            ("Classify each opportunity as device / database / software / application integration — this is the K1 middleware-type vocabulary.", ""),
            ("Score each candidate 1-5 for volume, error rate and business impact; rank the register by total score.", ""),
        ],
        test="Your register lists at least six candidates, each naming a source system, a target system, a connector and a score. The top-ranked candidate should be the service-call intake — the one you automate in Lab 3.",
        stages=[('Open the environment', 'course env picker'),
                ('Read the manual process', 'KinetEco service calls'),
                ('Map the swimlane', 'mark every re-keying'),
                ('Browse the connectors', 'source and target'),
                ('Score and rank', 'volume · errors · impact')],
    ),
    dict(
        num=2, topic=1,
        title="Feasibility Scan and Middleware Selection",
        objective="A2 Perform feasibility scan and assessment to identify potential middleware to be used; K1 Types of middleware and their features.",
        desc=(
            "Take the top three opportunities from Lab 1 and run a structured feasibility scan against "
            "six criteria: connector availability, data readiness, licensing, security and governance, "
            "performance/delegation, and maintainability. Decide for each whether Power Platform is the "
            "right middleware — and be prepared to justify a 'no'."
        ),
        build="A Feasibility Assessment Matrix scoring three candidates against six criteria, with a documented Build / Defer / Reject decision and justification for each.",
        services="Power Platform admin center, connector catalogue, Microsoft licensing documentation",
        environment=True,
        steps=[
            ("For each of your top three candidates, confirm a standard connector exists — and note whether it is Standard or Premium, because Premium changes the licence needed.", "https://make.powerautomate.com/connectors/"),
            ("Assess data readiness: is the source structured, does it have a stable key, is it an Excel Table (Power Apps cannot bind to a loose range)?", ""),
            ("Assess security and governance: who owns the data, does the flow need a service account, does a DLP policy separate the connectors you intend to combine?", ""),
            ("Check the environment's DLP policies in the admin center and record any that would block your design.", "https://admin.powerplatform.microsoft.com/"),
            ("Assess performance: estimate row counts and check each connector's delegation support and throttling limits.", ""),
            ("Score each criterion 1-5, total the scores, and record a Build / Defer / Reject decision with a one-line justification.", ""),
            ("Present your top candidate to the class in 2 minutes: the opportunity, the middleware, and why it is feasible.", ""),
        ],
        test="Each of the three candidates has six scores, a total, and a decision with justification. At least one candidate should be Deferred or Rejected with a defensible reason — a scan where everything passes is not a scan.",
        stages=[('Confirm connectors exist', 'Standard or Premium'),
                ('Assess data readiness', 'Table? stable key?'),
                ('Check security and DLP', 'who owns the data'),
                ('Assess performance', 'rows and delegation'),
                ('Decide and justify', 'Build · Defer · Reject')],
    ),
]
