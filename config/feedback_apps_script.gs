/**
 * Data Ekam — feedback Apps Script backend.
 *
 * Receives each feedback submission from the site's feedback form and
 * appends it as a new row in a Google Sheet, so feedback from every
 * visitor lands in one shared place (not just the visitor's own browser).
 *
 * SETUP
 *  1. Create a new Google Sheet. Add a header row: Date | Website | Name | Message | Status
 *  2. Extensions > Apps Script. Delete any starter code and paste in this
 *     entire file.
 *  3. Deploy > New deployment > select type "Web app".
 *       - Execute as: Me
 *       - Who has access: Anyone
 *  4. Click Deploy, authorize it, then copy the Web App URL it gives you
 *     (it ends in /exec).
 *  5. Paste that URL as the value of FEEDBACK_ENDPOINT in the main HTML
 *     file's feedback-submission script.
 *
 * This deployment is specific to Data Ekam's own sheet — a different site
 * (e.g. BharatGaruda.com, IntiGuttu.com) needs its own separate deployment
 * of this same script pointed at its own sheet, never a shared one.
 *
 * The form on the page posts as application/x-www-form-urlencoded (not
 * JSON) specifically so the request stays a CORS "simple request" — Apps
 * Script web apps don't handle the pre-flight OPTIONS request a JSON POST
 * would trigger. doPost() below reads the fields from e.parameter to match.
 */

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var p = (e && e.parameter) ? e.parameter : {};

    sheet.appendRow([
      p.Date || '',
      p.Website || '',
      p.Name || '',
      p.Message || '',
      p.Status || 'Open'
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Lets you sanity-check the deployment URL directly in a browser — it should
// show a small JSON status message instead of an error page.
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Data Ekam feedback endpoint is live. POST feedback here.' }))
    .setMimeType(ContentService.MimeType.JSON);
}
