const SHEET_NAME = "Claims";
const HEADERS = ["item_id", "item_name", "status", "claimed_by", "claimed_at"];
const ITEMS = [
  ["01-fennel-seeds", "Fennel Seeds"],
  ["02-flax-seeds", "Flax Seeds"],
  ["03-ajwain", "Ajwain"],
  ["04-methi-seeds", "Methi Seeds"],
  ["05-kalonji", "Kalonji"],
  ["06-dalia-split", "Dalia Split"],
  ["07-badam-drink-mix", "Badam Drink Mix"],
  ["08-multi-colored-peppers-mini-sachets", "Multi-Colored Peppers (Mini Sachets)"],
  ["09-quick-rolled-oats", "Quick Rolled Oats"],
  ["10-panipuri", "Panipuri"],
  ["11-star-anise", "Star Anise"],
  ["12-multi-sprinkles", "Multi Sprinkles"],
  ["13-black-peppers", "Black Peppers"],
  ["14-cinnamon-sticks", "Cinnamon Sticks"],
  ["15-black-cardamom", "Black Cardamom"],
  ["16-green-cardamom", "Green Cardamom"],
  ["17-kala-jeera", "Kala Jeera"],
  ["18-wet-tamarind", "Wet Tamarind"],
  ["19-javantri-mace", "Javantri (Mace)"],
  ["20-biryani-leaves", "Biryani Leaves"],
  ["21-white-corn-flour", "White Corn Flour"],
  ["22-urad-daal", "Urad Daal"],
  ["23-baking-mix-1", "Baking Mix 1"],
  ["24-baking-mix-2", "Baking Mix 2"],
  ["25-triple-chocolate-fudge-cake-mix", "Triple Chocolate Fudge Cake Mix"],
  ["26-funfetti-cake-mix", "Funfetti Cake Mix"],
  ["27-elbow-macaroni", "Elbow Macaroni"],
  ["28-toast-rusks-bread-rounds", "Toast Rusks / Bread Rounds"],
  ["29-kabuli-chana", "Kabuli Chana"],
  ["30-white-peas-dried-peas", "White Peas / Dried Peas"],
  ["31-gluten-free-chapati-flour", "Gluten-Free Chapati Flour"],
  ["32-all-purpose-flour", "All Purpose Flour"],
  ["33-raisins-bag", "Raisins (Bag)"],
  ["34-sunflower-kernels", "Sunflower Kernels"],
  ["35-raisins-jar", "Raisins (Jar)"],
  ["36-poppy-seeds", "Poppy Seeds"],
  ["37-sesame-seeds", "Sesame Seeds"],
  ["38-baking-soda", "Baking Soda"],
  ["39-amchur-powder", "Amchur Powder"],
  ["40-peanuts", "Peanuts"],
  ["41-zaatar", "Zaatar"]
];

function doGet(e) {
  const action = lowerTrim_(e && e.parameter && e.parameter.action) || "list";
  if (action !== "list") {
    return jsonResponse_({ ok: false, message: "Unknown action." });
  }

  return jsonResponse_({
    ok: true,
    items: listItems_()
  });
}

function doPost(e) {
  const action = lowerTrim_(e && e.parameter && e.parameter.action);
  if (action !== "claim") {
    return jsonResponse_({ ok: false, message: "Unknown action." });
  }

  return jsonResponse_(claimItem_(e.parameter || {}));
}

function syncItemsToSheet() {
  const sheet = getSheet_();
  ensureHeaders_(sheet);

  const existingIds = new Set(listItems_().map(item => item.item_id));
  ITEMS.forEach(([itemId, itemName]) => {
    if (!existingIds.has(itemId)) {
      sheet.appendRow([itemId, itemName, "Available", "", ""]);
    }
  });
}

function listItems_() {
  const sheet = getSheet_();
  ensureHeaders_(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  return sheet
    .getRange(2, 1, lastRow - 1, HEADERS.length)
    .getValues()
    .filter(row => row[0])
    .map(rowToItem_);
}

function claimItem_(params) {
  const itemId = trimText_(params.item_id);
  const itemName = trimText_(params.item_name);
  const claimedBy = trimText_(params.claimed_by);

  if (!itemId || !claimedBy) {
    return { ok: false, message: "item_id and claimed_by are required." };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = getSheet_();
    ensureHeaders_(sheet);

    const items = listItems_();
    const existingIndex = items.findIndex(item => item.item_id === itemId);

    if (existingIndex >= 0) {
      const existingItem = items[existingIndex];
      if (lowerTrim_(existingItem.status) === "claimed") {
        return {
          ok: false,
          error: "already_claimed",
          message: "Item already claimed.",
          item: existingItem
        };
      }

      const claimedAt = new Date().toISOString();
      const updatedItem = {
        item_id: itemId,
        item_name: itemName || existingItem.item_name || itemId,
        status: "Claimed",
        claimed_by: claimedBy,
        claimed_at: claimedAt
      };

      sheet
        .getRange(existingIndex + 2, 1, 1, HEADERS.length)
        .setValues([[updatedItem.item_id, updatedItem.item_name, updatedItem.status, updatedItem.claimed_by, updatedItem.claimed_at]]);

      return { ok: true, item: updatedItem };
    }

    const claimedAt = new Date().toISOString();
    const newItem = {
      item_id: itemId,
      item_name: itemName || itemId,
      status: "Claimed",
      claimed_by: claimedBy,
      claimed_at: claimedAt
    };

    sheet.appendRow([newItem.item_id, newItem.item_name, newItem.status, newItem.claimed_by, newItem.claimed_at]);
    return { ok: true, item: newItem };
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  return sheet;
}

function ensureHeaders_(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  const currentHeaders = headerRange.getValues()[0];
  const needsHeaders = HEADERS.some((header, index) => currentHeaders[index] !== header);

  if (needsHeaders) {
    headerRange.setValues([HEADERS]);
  }
}

function rowToItem_(row) {
  return {
    item_id: String(row[0] || ""),
    item_name: String(row[1] || ""),
    status: String(row[2] || "Available"),
    claimed_by: String(row[3] || ""),
    claimed_at: row[4] instanceof Date ? row[4].toISOString() : String(row[4] || "")
  };
}

function trimText_(value) {
  return String(value || "").trim();
}

function lowerTrim_(value) {
  return String(value || "").trim().toLowerCase();
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
