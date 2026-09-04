# Lab 7 — Canvas App from Excel Data

**Course:** Applications Integration with Power Apps and Power Automate (TGS-2022015539)  
**Topic 3:** Power Apps  
**Maps to:** K3 Different types of platforms on which applications run; A4 Support API-level integration.

---

## What you are building

Create your first canvas app over the Excel table you built in Lab 4. Power Apps generates a three-screen browse/detail/edit app; you then read how each screen is wired so the generated app becomes something you can modify rather than a black box.

**Deliverable —** Canvas app 'Lab 7 - Canvas App from Excel (DO NOT DELETE)' — a working three-screen app bound to the ServiceCalls Excel table, saved and published.

**Tools —** Power Apps Studio, Excel Online (Business), OneDrive for Business

> **Environment.** Every lab in this course runs in the Power Platform
> environment **TGS-2022015539-Applications Integration with Power Apps and Power Automate**.
> Check the environment picker in the top-right of the maker portal BEFORE you
> build anything — work created in the wrong environment cannot simply be moved.

> **Name the app exactly** `Lab 7 - Canvas App from Excel (DO NOT DELETE)`.

## Data

- `data/KinetEco Service Calls.xlsx` — upload this to your OneDrive for Business before you start.

Each workbook already contains a real Excel **Table**. Power Apps and Power
Automate can only bind to a named table, never to a loose range — if you build
your own workbook later, remember to Insert > Table first.

## Steps

1. Open make.powerapps.com and confirm the environment picker shows the course environment.
   - <https://make.powerapps.com>

2. Choose Start with data > Excel, connect to OneDrive for Business, and select the ServiceCalls table from KinetEco Service Calls.xlsx.

3. Let Power Apps generate the app, then immediately File > Save as with the name 'Lab 7 - Canvas App from Excel (DO NOT DELETE)'.

4. Select BrowseGallery1 and read its Items property — this is the formula that binds the gallery to the data source.

5. Change the Items formula to sort and search the table, then confirm the gallery reorders live.

   ```
   SortByColumns(Search(ServiceCalls, TextSearchBox1.Text, "Problem"), "Date_x0020_Reported", Descending)
   ```

6. Select the detail screen and inspect the Item property to see how the selected record is passed between screens.

   ```
   BrowseGallery1.Selected
   ```

7. On the edit screen, read the OnSuccess and OnFailure of the form, and the SubmitForm call on the tick icon.

   ```
   SubmitForm(EditForm1)
   ```

8. Preview with F5, add a record, and confirm it appears in the Excel workbook.

9. Publish the app, then note the Version history under Details — every publish is a restorable version.

## Test it

The app browses, filters and sorts real service-call data; adding a record through the app writes a new row to the Excel table, and the app appears in the Apps list as published.

---

◀ Previous: [Lab 6 — Leave Application Approval](../lab-06/lab-06.md)  
▶ Next: [Lab 8 — Blank Canvas App with Power Fx](../lab-08/lab-08.md)
