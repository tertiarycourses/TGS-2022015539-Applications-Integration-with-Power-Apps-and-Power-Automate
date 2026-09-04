"""Topic 3 — Power Apps (K3, K5, A4, A6).

Labs 7-10 move from flows to apps: generate an app from data, then build one from a
blank canvas with real formulas, then reach an external API through a custom connector,
then verify the app across platforms. Each lab consumes the data store built in Topic 2.
"""

DOMAIN3 = [
    dict(
        num=7, topic=3,
        title="Canvas App from Excel Data",
        objective="K3 Different types of platforms on which applications run; A4 Support API-level integration.",
        desc=(
            "Create your first canvas app over the Excel table you built in Lab 4. Power Apps generates "
            "a three-screen browse/detail/edit app; you then read how each screen is wired so the "
            "generated app becomes something you can modify rather than a black box."
        ),
        build="Canvas app 'Lab 7 - Canvas App from Excel (DO NOT DELETE)' — a working three-screen app bound to the ServiceCalls Excel table, saved and published.",
        services="Power Apps Studio, Excel Online (Business), OneDrive for Business",
        app_name="Lab 7 - Canvas App from Excel (DO NOT DELETE)",
        environment=True,
        steps=[
            ("Open make.powerapps.com and confirm the environment picker shows the course environment.", "https://make.powerapps.com"),
            ("Choose Start with data > Excel, connect to OneDrive for Business, and select the ServiceCalls table from KinetEco Service Calls.xlsx.", ""),
            ("Let Power Apps generate the app, then immediately File > Save as with the name 'Lab 7 - Canvas App from Excel (DO NOT DELETE)'.", ""),
            ("Select BrowseGallery1 and read its Items property — this is the formula that binds the gallery to the data source.", ""),
            ("Change the Items formula to sort and search the table, then confirm the gallery reorders live.", "SortByColumns(Search(ServiceCalls, TextSearchBox1.Text, \"Problem\"), \"Date_x0020_Reported\", Descending)"),
            ("Select the detail screen and inspect the Item property to see how the selected record is passed between screens.", "BrowseGallery1.Selected"),
            ("On the edit screen, read the OnSuccess and OnFailure of the form, and the SubmitForm call on the tick icon.", "SubmitForm(EditForm1)"),
            ("Preview with F5, add a record, and confirm it appears in the Excel workbook.", ""),
            ("Publish the app, then note the Version history under Details — every publish is a restorable version.", ""),
        ],
        test="The app browses, filters and sorts real service-call data; adding a record through the app writes a new row to the Excel table, and the app appears in the Apps list as published.",
        stages=[('Open the maker portal', 'check the environment'),
                ('Start with Excel data', 'pick the table'),
                ('Save with the lab name', 'three screens generated'),
                ('Read the Items formula', 'how binding works'),
                ('Sort, search and publish', 'then check versions')],
    ),
    dict(
        num=8, topic=3,
        title="Blank Canvas App with Power Fx",
        objective="K3 Different types of platforms on which applications run; A6 Verify proper functioning of modules and applications across multiple or integrated platforms.",
        desc=(
            "A generated app hides the mechanics. Build one from an empty screen so every control, "
            "variable and formula is yours: a leave-request form with validation, variables and "
            "navigation, which becomes the front end you wire to a flow in Topic 4."
        ),
        build="Canvas app 'Lab 8 - Leave Request App (DO NOT DELETE)' — a two-screen app built from blank with validated inputs, variables and navigation.",
        services="Power Apps Studio, Power Fx",
        app_name="Lab 8 - Leave Request App (DO NOT DELETE)",
        environment=True,
        steps=[
            ("Create a blank canvas app with Tablet layout named 'Lab 8 - Leave Request App (DO NOT DELETE)'.", ""),
            ("On Screen1 insert a Label as the title, then Text input controls named txtApplicant and txtReason, a Dropdown ddLeaveType, a Date picker dpStart and a Text input txtDays.", ""),
            ("Set the dropdown Items to a literal table of leave types.", "[\"Annual\",\"Medical\",\"Childcare\",\"Unpaid\"]"),
            ("Insert a Button named btnSubmit and set its DisplayMode so it is disabled until the form is complete — validation before submission.", "If(IsBlank(txtApplicant.Text) || IsBlank(txtDays.Text), DisplayMode.Disabled, DisplayMode.Edit)"),
            ("On btnSubmit OnSelect, set a global variable capturing the request and navigate to a confirmation screen.", "Set(gblRequest, {Applicant: txtApplicant.Text, LeaveType: ddLeaveType.Selected.Value, Start: dpStart.SelectedDate, Days: Value(txtDays.Text), Reason: txtReason.Text}); Navigate(scrConfirm, ScreenTransition.Cover)"),
            ("Add a second screen scrConfirm with labels reading back the variable, proving state carried across screens.", "\"Thank you \" & gblRequest.Applicant & \" - \" & gblRequest.Days & \" day(s) of \" & gblRequest.LeaveType"),
            ("Add a Back button on the confirmation screen that clears the variable and returns to the form.", "Set(gblRequest, Blank()); Reset(txtApplicant); Navigate(Screen1)"),
            ("Use the App checker to clear all formula errors and accessibility warnings.", ""),
            ("Save and publish the app.", ""),
        ],
        test="The submit button stays disabled until the required fields are filled; submitting navigates to the confirmation screen showing the entered values; App checker reports no errors.",
        stages=[('Create a blank canvas app', 'tablet layout'),
                ('Insert the form controls', 'inputs and dropdown'),
                ('Validate before submit', 'disable until complete'),
                ('Set a global variable', 'carry state across screens'),
                ('Navigate and confirm', 'read the values back')],
    ),
    dict(
        num=9, topic=3,
        title="Custom Connector and API Integration",
        objective="K5 Functions of Application Programming Interfaces (APIs); A4 Support API-level integration.",
        desc=(
            "Connectors are the platform's API layer. Where no prebuilt connector exists you build a "
            "custom one. Create a custom connector over a public REST API, define its action and "
            "response schema, test it, and consume it from a canvas app."
        ),
        build="Custom connector 'Lab 9 - Public Holiday API (DO NOT DELETE)' plus canvas app 'Lab 9 - API Consumer App (DO NOT DELETE)' displaying live API data.",
        services="Power Apps custom connectors, REST, JSON, Power Fx",
        app_name="Lab 9 - API Consumer App (DO NOT DELETE)",
        environment=True,
        steps=[
            ("In make.powerapps.com go to More > Discover all > Custom connectors, and choose New custom connector > Create from blank.", ""),
            ("Name it 'Lab 9 - Public Holiday API (DO NOT DELETE)'. On General set Host to date.nager.at and the Base URL to /api/v3.", ""),
            ("On Security choose No authentication — this API is public. Note in your workbook what you would choose for a real corporate API (OAuth 2.0 or API key).", ""),
            ("On Definition add a New action: Summary 'Get public holidays', Operation ID GetPublicHolidays.", ""),
            ("Set the Request by using Import from sample: verb GET and the URL below, which makes the year and country code path parameters.", "https://date.nager.at/api/v3/PublicHolidays/2026/SG"),
            ("Use Import from sample on the Response with a real payload so the connector learns the schema — this is what gives you typed dynamic content downstream.", ""),
            ("Create the connector, then open the Test tab, create a connection, and run the operation with year 2026 and country SG.", ""),
            ("Create a canvas app named 'Lab 9 - API Consumer App (DO NOT DELETE)', add the custom connector as a data source, and load the API result into a collection on a button.", "ClearCollect(colHolidays, 'Lab9-PublicHolidayAPI'.GetPublicHolidays(2026,\"SG\"))"),
            ("Add a gallery bound to colHolidays showing the date and the local name of each holiday.", ""),
        ],
        test="The connector Test tab returns HTTP 200 with a JSON array of holidays. In the app, pressing the button fills the gallery with live public-holiday data from the external API.",
        stages=[('Create a custom connector', 'from blank'),
                ('Set host and base URL', 'date.nager.at'),
                ('Choose the security', 'none for a public API'),
                ('Import the schema', 'from a sample response'),
                ('Consume it from an app', 'collect and display')],
    ),
    dict(
        num=10, topic=3,
        title="Cross-Platform Verification and Sharing",
        objective="A6 Verify proper functioning of modules and applications across multiple or integrated platforms; K3 Different types of platforms on which applications run.",
        desc=(
            "An app that only works on the maker's screen is not integrated. Verify your apps run "
            "correctly in the browser, on the mobile player and embedded in Teams, then share them "
            "with the right permissions and record the results in a test matrix."
        ),
        build="A completed Cross-Platform Verification Matrix for Labs 7-9, plus the apps shared and one app embedded in Microsoft Teams.",
        services="Power Apps mobile, Microsoft Teams, Power Apps sharing",
        environment=True,
        steps=[
            ("Build a test matrix with your apps as rows and Browser, Mobile and Teams as columns, plus a column for defects found.", ""),
            ("Run Lab 7 and Lab 8 in the browser at make.powerapps.com and record load time and any layout problem.", ""),
            ("Install Power Apps mobile, sign in with the same account, and open each app. Note where a Tablet-layout app is awkward on a phone — a real compatibility finding.", ""),
            ("In Teams, add the Power Apps app, then Add an app to a channel tab and select Lab 8. Confirm it renders and functions inside Teams.", ""),
            ("Share Lab 8 with a classmate as a User (not Co-owner) and have them confirm they can run but not edit it.", ""),
            ("Confirm your classmate also needs access to the underlying data source — sharing an app does not share its data. Record this as a finding.", ""),
            ("Complete the matrix with a pass/fail and at least three concrete defects or limitations across the platforms.", ""),
        ],
        test="Every cell of the matrix is filled for all three platforms, with at least three real defects or limitations recorded, and a classmate can run your shared app.",
        stages=[('Build a test matrix', 'apps × platforms'),
                ('Test in the browser', 'load time and layout'),
                ('Test on mobile', 'note what is awkward'),
                ('Embed in Teams', 'add as a channel tab'),
                ('Share and record findings', 'User, not Co-owner')],
    ),
]
