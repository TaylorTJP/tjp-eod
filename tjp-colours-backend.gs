/************************************************************************
 * TJP Colour Selection - Google Apps Script backend
 *
 * What it does when a client submits colours from tjp-colours.html:
 *   1. Finds (or creates) that job's folder in your Google Drive.
 *   2. Saves the colour selection as a tidy text file inside it.
 *   3. Emails you a notification with the colours.
 *
 * It also answers ?jobs=1 with your live job list (from a Jobs sheet),
 * so the colour page's preview picker matches your real jobs/folders.
 *
 * ------------------------- SETUP (one time) -------------------------
 * 1. In Google Drive, note the folder that holds your per-job folders
 *    (the same "back-costing" jobs folder). Open it and copy the ID from
 *    the URL: drive.google.com/drive/folders/THIS_IS_THE_ID
 * 2. Paste that ID into JOBS_PARENT_FOLDER_ID below.
 * 3. Set NOTIFY_EMAIL to taylor@tjphobart.com.au.
 * 4. (Optional) If you keep a Jobs list in a Google Sheet, put its ID in
 *    JOBS_SHEET_ID and the tab name in JOBS_SHEET_TAB. Leave blank to skip.
 * 5. Deploy: Deploy > New deployment > type "Web app" >
 *      Execute as: Me  |  Who has access: Anyone
 *    Copy the /exec URL it gives you.
 * 6. Paste that /exec URL into API_URL at the top of tjp-colours.html.
 * That's it - the page goes from demo to live.
 ********************************************************************/

// ---- CONFIG ----
const JOBS_PARENT_FOLDER_ID = "PASTE_PARENT_FOLDER_ID"; // folder that contains each job's folder
const NOTIFY_EMAIL          = "taylor@tjphobart.com.au";
const JOBS_SHEET_ID         = ""; // optional: a Google Sheet ID holding the job list
const JOBS_SHEET_TAB        = "Jobs"; // tab name in that sheet (column A = job names)

// ---- POST: a client submitted their colours ----
function doPost(e){
  try{
    const rec = JSON.parse(e.postData.contents);
    // Folder is named by the property address (fallback: client name).
    const folder = getJobFolder(rec.client_address || rec.client_name);
    const text = formatSelection(rec);

    // Save a text file into the job folder
    const fname = "Colour selection - " + (rec.date || todayStr()) + " - " + safe(rec.client_name || "client") + ".txt";
    folder.createFile(fname, text, MimeType.PLAIN_TEXT);

    // Notify you
    const subject = "🎨 Colours in: " + (rec.client_address || rec.client_name || "client");
    MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: subject, body: text });

    return json({ ok: true });
  }catch(err){
    return json({ ok: false, error: String(err) });
  }
}

// ---- GET: return the job list for the page's preview picker ----
function doGet(e){
  if (e && e.parameter && e.parameter.jobs){
    return json(getJobList());
  }
  return ContentService.createTextOutput("TJP colours backend OK");
}

// ---- helpers ----
function getJobFolder(jobName){
  const parent = DriveApp.getFolderById(JOBS_PARENT_FOLDER_ID);
  const wanted = (jobName || "Unassigned").trim();
  const it = parent.getFoldersByName(wanted);
  if (it.hasNext()) return it.next();
  return parent.createFolder(wanted); // create if it doesn't exist yet
}

function formatSelection(rec){
  let out = "";
  out += "TJP COLOUR SELECTION\n";
  out += "====================\n\n";
  out += "Client:   " + (rec.client_name || "") + "\n";
  out += "Address:  " + (rec.client_address || "") + "\n";
  out += "Contact:  " + (rec.client_contact || "") + "\n";
  out += "Sent:     " + (rec.date || "") + " " + (rec.time || "") + "\n";
  out += "Ref:      " + (rec.docket || "") + "\n";
  writeScope(rec, "Interior");
  writeScope(rec, "Exterior");
  out += "\n---\nColours are indicative; confirm with a physical sample before ordering.\n";
  return out;

  function writeScope(rec, scope){
    const items = (rec.items || []).filter(function(it){ return it.scope === scope; });
    if (!items.length) return;
    out += "\n" + scope.toUpperCase() + "\n" + "-".repeat(scope.length) + "\n";
    items.forEach(function(it){
      out += "\n• " + (it.area || "(area)") + "\n";
      out += "    Colour: " + (it.colour || "") + (it.code ? ("  (Dulux " + it.code + ")") : "") + (it.custom ? "  [client-entered]" : "") + "\n";
      if (it.finish) out += "    Finish: " + it.finish + "\n";
      if (it.note)   out += "    Note:   " + it.note + "\n";
    });
  }
}

function getJobList(){
  if (!JOBS_SHEET_ID) return [];
  try{
    const sh = SpreadsheetApp.openById(JOBS_SHEET_ID).getSheetByName(JOBS_SHEET_TAB);
    const vals = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
    return vals.map(function(r){ return String(r[0]).trim(); }).filter(function(v){ return v && v.toLowerCase() !== "job" && v.toLowerCase() !== "jobs"; });
  }catch(err){ return []; }
}

function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function safe(s){ return String(s).replace(/[\\/:*?"<>|]+/g, "-").slice(0, 60); }
function todayStr(){ return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"); }
