# Lab 10 — Cross-Platform Verification and Sharing

**Course:** Applications Integration with Power Apps and Power Automate (TGS-2022015539)  
**Topic 3:** Power Apps  
**Maps to:** A6 Verify proper functioning of modules and applications across multiple or integrated platforms; K3 Different types of platforms on which applications run.

---

## What you are building

An app that only works on the maker's screen is not integrated. Verify your apps run correctly in the browser, on the mobile player and embedded in Teams, then share them with the right permissions and record the results in a test matrix.

**Deliverable —** A completed Cross-Platform Verification Matrix for Labs 7-9, plus the apps shared and one app embedded in Microsoft Teams.

**Tools —** Power Apps mobile, Microsoft Teams, Power Apps sharing

> **Environment.** Every lab in this course runs in the Power Platform
> environment **TGS-2022015539-Applications Integration with Power Apps and Power Automate**.
> Check the environment picker in the top-right of the maker portal BEFORE you
> build anything — work created in the wrong environment cannot simply be moved.

## Data

- `data/KinetEco Service Calls.xlsx` — upload this to your OneDrive for Business before you start.

Each workbook already contains a real Excel **Table**. Power Apps and Power
Automate can only bind to a named table, never to a loose range — if you build
your own workbook later, remember to Insert > Table first.

## Steps

1. Build a test matrix with your apps as rows and Browser, Mobile and Teams as columns, plus a column for defects found.

2. Run Lab 7 and Lab 8 in the browser at make.powerapps.com and record load time and any layout problem.

3. Install Power Apps mobile, sign in with the same account, and open each app. Note where a Tablet-layout app is awkward on a phone — a real compatibility finding.

4. In Teams, add the Power Apps app, then Add an app to a channel tab and select Lab 8. Confirm it renders and functions inside Teams.

5. Share Lab 8 with a classmate as a User (not Co-owner) and have them confirm they can run but not edit it.

6. Confirm your classmate also needs access to the underlying data source — sharing an app does not share its data. Record this as a finding.

7. Complete the matrix with a pass/fail and at least three concrete defects or limitations across the platforms.

## Test it

Every cell of the matrix is filled for all three platforms, with at least three real defects or limitations recorded, and a classmate can run your shared app.

---

◀ Previous: [Lab 9 — Custom Connector and API Integration](../lab-09/lab-09.md)  
▶ Next: [Lab 11 — Call a Flow from a Canvas App](../lab-11/lab-11.md)
