# Rowe Split — Google Sheets 2-Way Sync Setup

Connect your **Rowe Split** rent tracker to Google Sheets so you, Zong, and your family always see the exact same live ledger data across all your phones and computers.

---

### Step 1: Create a Google Sheet
1. Open [sheet.new](https://sheet.new) in your browser.
2. Name the sheet: **Rowe Split Ledger**.

---

### Step 2: Add the Google Apps Script
1. In the Google Sheet, click **Extensions** > **Apps Script**.
2. Delete any existing code in the editor (`function myFunction() { ... }`).
3. Paste the following script:

```javascript
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("LedgerData") || ss.getSheets()[0];
  var val = sheet.getRange("A1").getValue();
  return ContentService.createTextOutput(val || "{}").setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheetByName("LedgerData");
    if (!dataSheet) dataSheet = ss.insertSheet("LedgerData");
    var payload = e.postData.contents;
    dataSheet.getRange("A1").setValue(payload);
    updateSummarySheet(ss, JSON.parse(payload));
    return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function updateSummarySheet(ss, state) {
  var sheet = ss.getSheetByName("Summary");
  if (!sheet) sheet = ss.insertSheet("Summary", 0);
  sheet.clear();
  sheet.appendRow(["Month", "Description", "Category", "Amount", "Plato Share ($)", "Zong Share ($)", "Plato Paid", "Zong Paid", "Plato Comments", "Zong Comments"]);
  sheet.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#e8f0fe");
  var months = state.months || {};
  for (var id in months) {
    var m = months[id];
    var items = m.lineItems || [];
    if (items.length === 0) {
      sheet.appendRow([id, "(Blank)", "-", 0, 0, 0, m.paid && m.paid.plato ? "Yes" : "No", m.paid && m.paid.zong ? "Yes" : "No", m.notesPlato || "", m.notesZong || ""]);
    } else {
      for (var i = 0; i < items.length; i++) {
        var it = items[i];
        var p = (it.amount * it.splitPlato) / 100;
        sheet.appendRow([id, it.name, it.category, it.amount, p, it.amount - p, m.paid && m.paid.plato ? "Yes" : "No", m.paid && m.paid.zong ? "Yes" : "No", i === 0 ? (m.notesPlato || "") : "", i === 0 ? (m.notesZong || "") : ""]);
      }
    }
  }
}
```

4. Click **Save** (💾 disk icon).

---

### Step 3: Deploy as Web App
1. In the top right of Apps Script, click **Deploy** > **New deployment**.
2. Click the **gear icon (Select type)** and choose **Web app**.
3. Set the following settings:
   - **Description**: `Rowe Split API`
   - **Execute as**: `Me (<your email>)`
   - **Who has access**: `Anyone` *(Crucial so both your and Zong's phones can sync)*
4. Click **Deploy**.
5. Grant permissions if prompted (Click *Advanced* > *Go to Rowe Split (unsafe)*).
6. Copy the **Web App URL** (it looks like `https://script.google.com/macros/s/.../exec`).

---

### Step 4: Connect in Rowe Split
1. Open [index.html](index.html).
2. Click **Connect Google Sheet** in the top header.
3. Paste your Web App URL and click **Save & Sync Now**.

---

### How Sharing Works Once Connected
- **Share (View-Only)**: Click **Share (View-Only)** in the header. Anyone opening this link (family, friends, or roommates without editing access) will see the live numbers in **100% Non-Editing / View-Only Mode** (no editable boxes, inputs locked, toggles disabled).
- **Two-Way Sync**: Anyone with access to the Google Sheet URL can paste it into **Connect Google Sheet** to collaborate with full editing capabilities.
