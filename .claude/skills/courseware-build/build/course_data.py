"""
SINGLE SOURCE OF TRUTH — Applications Integration with Power Apps and Power Automate
TGS-2022015539.  Every artifact (PPT, LP, LG, LG.md, labs index, assessment) is
generated from this module + data_domain1..4.py so they stay 100% aligned.

Accredited spec (WSQ_CP_TIPL_power_automate_v2_1.docx):
  TSC  : Applications Integration (ICT-DIT-3003-1.1), Proficiency Level 3
  Hours: Classroom 7.5h + Practical 6h = 13.5h delivery, + 2.5h assessment = 16h
  Assess: WA(SAQ) 1h  +  PP 1.5h
"""

# ------------------------------------------------------------------ metadata
TITLE        = "Applications Integration with Power Apps and Power Automate"
SHORT_TITLE  = "Applications Integration with Power Apps and Power Automate"
COURSE_CODE  = "TGS-2022015539"
VERSION      = "v8.0"
VERSION_DATE = "4 September 2026"
ORG          = "Tertiary Infotech Academy Pte Ltd"
UEN          = "UEN: 201200696W"
TRAINER      = "Dr Alfred Ang"
DAYS         = 2

TSC_TITLE    = "Applications Integration"
TSC_CODE     = "ICT-DIT-3003-1.1"
TSC_LEVEL    = "Level 3"

COURSE_URL   = "https://www.tertiarycourses.com.sg/wsq-applications-integration-with-power-apps-and-power-automate.html"

# Power Platform environment provisioned for the hands-on labs
LAB_ENVIRONMENT = "TGS-2022015539-Applications Integration with Power Apps and Power Automate"

# ------------------------------------------------------------------ outcomes
LEARNING_OUTCOMES = [
    "LO1: Identify opportunities to use Power Platform apps and perform a feasibility scan for potential application integration",
    "LO2: Utilise and test Power Automate flows for business automation",
    "LO3: Create Power Apps canvas apps and verify their functionalities",
    "LO4: Improve Power Apps with Power Automate flows, and resolve integration issues",
]

# TSC knowledge (K) and ability (A) statements — assessment must cover EVERY one.
TSC_KNOWLEDGE = [
    ("K1", "Types of middleware and their features"),
    ("K2", "Proper usage of middleware"),
    ("K3", "Different types of platforms on which applications run"),
    ("K4", "Potential technical, compatibility or performance issues in application integration"),
    ("K5", "Functions of Application Programming Interfaces (APIs)"),
]

TSC_ABILITIES = [
    ("A1", "Identify opportunities for creating connections among various devices, databases, software and applications"),
    ("A2", "Perform feasibility scan and assessment to identify potential middleware to be used"),
    ("A3", "Utilise middleware to integrate data and functions across application programs within an enterprise"),
    ("A4", "Support API-level integration"),
    ("A5", "Perform tests and checks on the connections between disparate application programs"),
    ("A6", "Verify proper functioning of modules and applications across multiple or integrated platforms"),
    ("A7", "Highlight technical, compatibility or performance issues following integration of applications or platforms on which they are used"),
    ("A8", "Implement modifications to mitigate the issues identified"),
]

# ------------------------------------------------------------------ topics (= learning units)
TOPICS = [
    dict(num=1, code="01",
         title="Opportunities for Using Power Platform Apps",
         subtitle="Integration landscape · feasibility scan · Power Platform product family",
         weighting="K1, A1, A2",
         mapping="ELO1 (K1, A1, A2)",
         concepts=[
            "Application integration connects disparate systems so data and functions flow between them without manual re-keying.",
            "Middleware is the software layer that sits between applications and brokers that exchange — Power Platform is low-code integration middleware.",
            "The Power Platform family: Power Apps (build), Power Automate (automate), Power BI (analyse), Copilot Studio (converse), Dataverse (store).",
            "A feasibility scan weighs business value, data readiness, connector availability, licensing, security and governance before you build.",
            "Connectors are the API surface of the platform: 1,000+ prebuilt connectors plus custom connectors for any REST API.",
         ]),
    dict(num=2, code="02",
         title="Power Automate",
         subtitle="Cloud flows · triggers and actions · conditions · approvals · testing",
         weighting="K2, A3, A5",
         mapping="ELO2 (K2, A3, A5)",
         concepts=[
            "A cloud flow is a trigger plus a sequence of actions; the trigger is the event that starts the flow.",
            "Flow types: automated (event-driven), instant (manual/button), scheduled (recurrence), and desktop (RPA).",
            "Dynamic content passes output from one action into the input of the next — this is how data crosses application boundaries.",
            "Conditions, loops and scopes give a flow its control logic; expressions add computation.",
            "Testing is a first-class step: Flow checker, Test, and the 28-day run history are how you prove a connection works.",
         ]),
    dict(num=3, code="03",
         title="Power Apps",
         subtitle="Canvas apps · data sources · controls and formulas · publishing and sharing",
         weighting="K3, K5, A4, A6",
         mapping="ELO3 (K3, K5, A4, A6)",
         concepts=[
            "Canvas apps give pixel-level control of the UI; model-driven apps generate the UI from the data model.",
            "Apps run across platforms — browser, iOS, Android, Teams and embedded in SharePoint or Power BI.",
            "A data source is reached through a connector; the same app can bind to Dataverse, Excel, SharePoint or a REST API.",
            "Power Fx is the Excel-like formula language that binds controls to data and behaviour.",
            "Publish, version and share are the app lifecycle: every save creates a version you can restore.",
         ]),
    dict(num=4, code="04",
         title="Integrate Power Apps and Power Automate",
         subtitle="Calling flows from apps · approvals · returning data · troubleshooting integration",
         weighting="K4, A7, A8",
         mapping="ELO4 (K4, A7, A8)",
         concepts=[
            "A canvas app calls a flow with the Power Apps (V2) trigger; parameters are typed and passed from controls.",
            "Respond to a Power App or flow returns data back to the app, closing the round trip.",
            "Approvals combine a flow, an approver and a decision branch — the classic human-in-the-loop integration.",
            "Integration issues cluster into technical (auth, connector), compatibility (data type, format) and performance (delegation, throughput).",
            "Delegation is the single most common Power Apps performance defect — know the warning and the mitigations.",
         ]),
]

# ------------------------------------------------------------------ day themes
DAY_THEMES = {
    1: "Opportunities, feasibility and building automation with Power Automate",
    2: "Building canvas apps, integrating them with flows, and assessment",
}

# ------------------------------------------------------------------ assessment
ASSESSMENT = dict(
    written="Written Assessment (WA) — Short-Answer Questions (SAQ), 1 hour, individual, open book. 5 questions covering K1–K5.",
    practical="Practical Performance (PP) — hands-on scenario tasks, 1.5 hours, individual, open book. 4 tasks covering A1–A8.",
    note="A minimum of 75% attendance is required to be eligible for assessment and funding. Learners must be assessed Competent in both WA and PP.",
)

# ------------------------------------------------------------------ recommended courses
RECOMMENDED_COURSES = [
    "WSQ - Business Process Automation with Power Automate and Copilot Studio Agents",
    "WSQ - Data Analytics and Visualisation with Power BI",
    "WSQ - Microsoft SharePoint for Collaboration",
    "WSQ - Data Visualisation with Tableau",
    "WSQ - Generative AI for Business Productivity",
]
