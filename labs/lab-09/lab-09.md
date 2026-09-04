# Lab 9 — Custom Connector and API Integration

**Course:** Applications Integration with Power Apps and Power Automate (TGS-2022015539)  
**Topic 3:** Power Apps  
**Maps to:** K5 Functions of Application Programming Interfaces (APIs); A4 Support API-level integration.

---

## What you are building

Connectors are the platform's API layer. Where no prebuilt connector exists you build a custom one. Create a custom connector over a public REST API, define its action and response schema, test it, and consume it from a canvas app.

**Deliverable —** Custom connector 'Lab 9 - Public Holiday API (DO NOT DELETE)' plus canvas app 'Lab 9 - API Consumer App (DO NOT DELETE)' displaying live API data.

**Tools —** Power Apps custom connectors, REST, JSON, Power Fx

> **Environment.** Every lab in this course runs in the Power Platform
> environment **TGS-2022015539-Applications Integration with Power Apps and Power Automate**.
> Check the environment picker in the top-right of the maker portal BEFORE you
> build anything — work created in the wrong environment cannot simply be moved.

> **Name the app exactly** `Lab 9 - API Consumer App (DO NOT DELETE)`.

## Data

- `data/Employee Survey.xlsx` — upload this to your OneDrive for Business before you start.

Each workbook already contains a real Excel **Table**. Power Apps and Power
Automate can only bind to a named table, never to a loose range — if you build
your own workbook later, remember to Insert > Table first.

## Steps

1. In make.powerapps.com go to More > Discover all > Custom connectors, and choose New custom connector > Create from blank.

2. Name it 'Lab 9 - Public Holiday API (DO NOT DELETE)'. On General set Host to date.nager.at and the Base URL to /api/v3.

3. On Security choose No authentication — this API is public. Note in your workbook what you would choose for a real corporate API (OAuth 2.0 or API key).

4. On Definition add a New action: Summary 'Get public holidays', Operation ID GetPublicHolidays.

5. Set the Request by using Import from sample: verb GET and the URL below, which makes the year and country code path parameters.
   - <https://date.nager.at/api/v3/PublicHolidays/2026/SG>

6. Use Import from sample on the Response with a real payload so the connector learns the schema — this is what gives you typed dynamic content downstream.

7. Create the connector, then open the Test tab, create a connection, and run the operation with year 2026 and country SG.

8. Create a canvas app named 'Lab 9 - API Consumer App (DO NOT DELETE)', add the custom connector as a data source, and load the API result into a collection on a button.

   ```
   ClearCollect(colHolidays, 'Lab9-PublicHolidayAPI'.GetPublicHolidays(2026,"SG"))
   ```

9. Add a gallery bound to colHolidays showing the date and the local name of each holiday.

## Test it

The connector Test tab returns HTTP 200 with a JSON array of holidays. In the app, pressing the button fills the gallery with live public-holiday data from the external API.

---

◀ Previous: [Lab 8 — Blank Canvas App with Power Fx](../lab-08/lab-08.md)  
▶ Next: [Lab 10 — Cross-Platform Verification and Sharing](../lab-10/lab-10.md)
