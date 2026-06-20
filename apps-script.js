// ============================================================
// ArcheoVault — Google Apps Script Backend
// Paste this into: Extensions → Apps Script in your Google Sheet
// Deploy as: Web App → Anyone with the link
// ============================================================

const CATEGORIES_SHEET = 'Κατηγορίες';
const ITEMS_SHEET = 'Αντικείμενα';
const EDITORS_SHEET = 'Editors';

// ── Auth Helpers ────────────────────────────────────────────

function getEditorEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(EDITORS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(EDITORS_SHEET);
    sheet.appendRow(['email']);
    // Seed with the spreadsheet owner
    const owner = ss.getOwner().getEmail();
    sheet.appendRow([owner]);
  }
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(r => String(r[0]).toLowerCase().trim()).filter(Boolean);
}

function verifyIdToken(idToken) {
  if (!idToken) return null;
  try {
    const payload = JSON.parse(
      Utilities.newBlob(
        Utilities.base64DecodeWebSafe(idToken.split('.')[1])
      ).getDataAsString()
    );
    // Check expiration
    const now = Math.floor(Date.now() / 1000);
    if (payload.exp && payload.exp < now) return null;
    // Check issuer
    if (payload.iss !== 'accounts.google.com' && payload.iss !== 'https://accounts.google.com') return null;
    return {
      email: (payload.email || '').toLowerCase(),
      name: payload.name || '',
      picture: payload.picture || ''
    };
  } catch (e) {
    return null;
  }
}

function isEditor(email) {
  if (!email) return false;
  const editors = getEditorEmails();
  return editors.includes(email.toLowerCase());
}

function requireEditor(body) {
  const token = body._idToken;
  if (!token) return { ok: false, error: 'AUTH_REQUIRED', message: 'Απαιτείται σύνδεση Google.' };
  const user = verifyIdToken(token);
  if (!user) return { ok: false, error: 'INVALID_TOKEN', message: 'Μη έγκυρο token. Συνδεθείτε ξανά.' };
  if (!isEditor(user.email)) return { ok: false, error: 'NOT_EDITOR', message: 'Δεν έχετε δικαίωμα επεξεργασίας (' + user.email + ').' };
  return { ok: true, user: user };
}

// ── Data Helpers ────────────────────────────────────────────

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === CATEGORIES_SHEET) {
      sheet.appendRow(['id', 'name']);
    } else if (name === ITEMS_SHEET) {
      sheet.appendRow(['id', 'name', 'category', 'location', 'description', 'photos', 'createdAt', 'updatedAt']);
    }
  }
  return sheet;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) return i + 1; // 1-based row number
  }
  return -1;
}

// ── GET handler (public — no auth required) ─────────────────

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';

  try {
    switch (action) {
      case 'getCategories':
        return jsonResponse({ success: true, categories: sheetToObjects(getSheet(CATEGORIES_SHEET)) });

      case 'getItems': {
        let items = sheetToObjects(getSheet(ITEMS_SHEET));
        const category = e.parameter.category;
        const search = (e.parameter.search || '').toLowerCase();
        if (category) {
          items = items.filter(it => it.category === category);
        }
        if (search) {
          items = items.filter(it =>
            (it.name || '').toLowerCase().includes(search) ||
            (it.category || '').toLowerCase().includes(search) ||
            (it.location || '').toLowerCase().includes(search) ||
            (it.description || '').toLowerCase().includes(search)
          );
        }
        return jsonResponse({ success: true, items: items });
      }

      case 'exportAll': {
        const categories = sheetToObjects(getSheet(CATEGORIES_SHEET));
        const items = sheetToObjects(getSheet(ITEMS_SHEET));
        return jsonResponse({ success: true, categories: categories, items: items });
      }

      case 'checkAccess': {
        const token = e.parameter.token;
        if (!token) return jsonResponse({ success: true, isEditor: false });
        const user = verifyIdToken(token);
        if (!user) return jsonResponse({ success: true, isEditor: false });
        return jsonResponse({ success: true, isEditor: isEditor(user.email), email: user.email, name: user.name });
      }

      default:
        return jsonResponse({ success: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ── POST handler (auth required) ────────────────────────────

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || '';

    // All POST actions require editor access
    const auth = requireEditor(body);
    if (!auth.ok) {
      return jsonResponse({ success: false, error: auth.error, message: auth.message });
    }

    switch (action) {

      // ── Categories ──────────────────────────────────────

      case 'addCategory': {
        const sheet = getSheet(CATEGORIES_SHEET);
        const id = Utilities.getUuid();
        sheet.appendRow([id, body.name]);
        return jsonResponse({ success: true, category: { id: id, name: body.name } });
      }

      case 'renameCategory': {
        const sheet = getSheet(CATEGORIES_SHEET);
        const row = findRowById(sheet, body.id);
        if (row === -1) return jsonResponse({ success: false, error: 'Category not found' });
        const oldName = sheet.getRange(row, 2).getValue();
        sheet.getRange(row, 2).setValue(body.name);
        const itemsSheet = getSheet(ITEMS_SHEET);
        const itemData = itemsSheet.getDataRange().getValues();
        for (let i = 1; i < itemData.length; i++) {
          if (itemData[i][2] === oldName) {
            itemsSheet.getRange(i + 1, 3).setValue(body.name);
          }
        }
        return jsonResponse({ success: true });
      }

      case 'deleteCategory': {
        const sheet = getSheet(CATEGORIES_SHEET);
        const row = findRowById(sheet, body.id);
        if (row === -1) return jsonResponse({ success: false, error: 'Category not found' });
        const catName = sheet.getRange(row, 2).getValue();
        const itemsSheet = getSheet(ITEMS_SHEET);
        const itemData = itemsSheet.getDataRange().getValues();
        const itemCount = itemData.slice(1).filter(r => r[2] === catName).length;
        if (itemCount > 0 && !body.force) {
          return jsonResponse({ success: false, error: 'HAS_ITEMS', itemCount: itemCount, categoryName: catName });
        }
        if (body.force && body.deleteItems) {
          for (let i = itemData.length - 1; i >= 1; i--) {
            if (itemData[i][2] === catName) {
              itemsSheet.deleteRow(i + 1);
            }
          }
        }
        if (body.force && body.reassignTo) {
          for (let i = 1; i < itemData.length; i++) {
            if (itemData[i][2] === catName) {
              itemsSheet.getRange(i + 1, 3).setValue(body.reassignTo);
              itemsSheet.getRange(i + 1, 8).setValue(new Date().toISOString());
            }
          }
        }
        sheet.deleteRow(row);
        return jsonResponse({ success: true });
      }

      // ── Items ───────────────────────────────────────────

      case 'addItem': {
        const sheet = getSheet(ITEMS_SHEET);
        const id = Utilities.getUuid();
        const now = new Date().toISOString();
        sheet.appendRow([
          id,
          body.name || '',
          body.category || '',
          body.location || '',
          body.description || '',
          body.photos || '',
          now,
          now
        ]);
        return jsonResponse({
          success: true,
          item: {
            id: id, name: body.name, category: body.category,
            location: body.location, description: body.description,
            photos: body.photos || '', createdAt: now, updatedAt: now
          }
        });
      }

      case 'updateItem': {
        const sheet = getSheet(ITEMS_SHEET);
        const row = findRowById(sheet, body.id);
        if (row === -1) return jsonResponse({ success: false, error: 'Item not found' });
        const now = new Date().toISOString();
        if (body.name !== undefined) sheet.getRange(row, 2).setValue(body.name);
        if (body.category !== undefined) sheet.getRange(row, 3).setValue(body.category);
        if (body.location !== undefined) sheet.getRange(row, 4).setValue(body.location);
        if (body.description !== undefined) sheet.getRange(row, 5).setValue(body.description);
        if (body.photos !== undefined) sheet.getRange(row, 6).setValue(body.photos);
        sheet.getRange(row, 8).setValue(now);
        return jsonResponse({ success: true, updatedAt: now });
      }

      case 'deleteItem': {
        const sheet = getSheet(ITEMS_SHEET);
        const row = findRowById(sheet, body.id);
        if (row === -1) return jsonResponse({ success: false, error: 'Item not found' });
        sheet.deleteRow(row);
        return jsonResponse({ success: true });
      }

      case 'moveItem': {
        const sheet = getSheet(ITEMS_SHEET);
        const row = findRowById(sheet, body.id);
        if (row === -1) return jsonResponse({ success: false, error: 'Item not found' });
        const now = new Date().toISOString();
        sheet.getRange(row, 4).setValue(body.location);
        sheet.getRange(row, 8).setValue(now);
        return jsonResponse({ success: true, updatedAt: now });
      }

      case 'importAll': {
        if (body.categories) {
          const catSheet = getSheet(CATEGORIES_SHEET);
          catSheet.clear();
          catSheet.appendRow(['id', 'name']);
          body.categories.forEach(c => catSheet.appendRow([c.id, c.name]));
        }
        if (body.items) {
          const itemSheet = getSheet(ITEMS_SHEET);
          itemSheet.clear();
          itemSheet.appendRow(['id', 'name', 'category', 'location', 'description', 'photos', 'createdAt', 'updatedAt']);
          body.items.forEach(it => itemSheet.appendRow([
            it.id, it.name, it.category, it.location,
            it.description, it.photos, it.createdAt, it.updatedAt
          ]));
        }
        return jsonResponse({ success: true });
      }

      default:
        return jsonResponse({ success: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}
