# How should Decade sound?

A team quiz for Decade's tone of voice: 24 moments in three parts, four options each. Answers go to a Google Sheet.

- `index.html` is the whole quiz. Picks stay in the browser until the person sends them.
- `apps-script.gs` is the Google Apps Script that writes each submission to the "Answers" tab of the sheet, one row per person. Sending again replaces that person's row.

## Connecting the sheet

1. Create a Google Sheet, then open Extensions > Apps Script.
2. Replace the sample code with `apps-script.gs` and save.
3. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Deploy and approve the permissions.
4. Copy the web app URL (it ends in `/exec`) and paste it into `SHEET_URL` near the top of the script in `index.html`.

If you change `apps-script.gs` later, use Deploy > Manage deployments > Edit > New version, so the URL stays the same.
