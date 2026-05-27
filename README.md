# Jin Diary

A static baby activity diary for GitHub Pages, with optional Google Sheets storage through Google Apps Script.

## Use with GitHub Pages

Upload these files to a GitHub repository and enable GitHub Pages. The app runs without a build step.

## Connect Google Sheets

1. Create a Google Sheet.
2. Open `Extensions > Apps Script`.
3. Paste the code from `google-apps-script.js`.
4. Deploy as a web app.
5. Set access to `Anyone`.
6. Paste the web app URL into the app's `Sheet` tab.

The app keeps a local copy in the browser and syncs with the sheet when the URL is configured.

Use the Apps Script deployment URL that looks like `https://script.google.com/macros/s/.../exec`. The redirected `https://script.googleusercontent.com/macros/echo?...` URL cannot receive the app's `append` and `delete` actions.

After changing `google-apps-script.js`, create a new Apps Script deployment version. Existing web app URLs keep running the old deployed code until you redeploy.

`Sync with sheet` uploads local-only entries first, then reloads the final data from the sheet.

The Sheet settings tab is visible only during local testing. It is hidden automatically on GitHub Pages.

## Background image

The page background is loaded from `assets/jin-background.jpg`.
