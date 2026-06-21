// ============================================================
// ArcheoVault — Google Apps Script Backend
// Paste this into: Extensions → Apps Script in your Google Sheet
// Deploy as: Web App → Anyone with the link
// ============================================================

const CATEGORIES_SHEET = 'Κατηγορίες';
const ITEMS_SHEET = 'Αντικείμενα';
const EDITORS_SHEET = 'Editors';
const TRANSLATIONS_SHEET = 'Translations';

// ── Auth Helpers ────────────────────────────────────────────

function getEditorEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(EDITORS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(EDITORS_SHEET);
    sheet.appendRow(['email']);
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
    const now = Math.floor(Date.now() / 1000);
    if (payload.exp && payload.exp < now) return null;
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
      sheet.appendRow(['id', 'name', 'parentId']);
    } else if (name === ITEMS_SHEET) {
      sheet.appendRow(['id', 'name', 'category', 'location', 'description', 'photos', 'createdAt', 'updatedAt']);
    }
  }
  return sheet;
}

function ensureCategoryParentIdColumn(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('parentId') === -1) {
    const col = headers.length + 1;
    sheet.getRange(1, col).setValue('parentId');
  }
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
    headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
    return obj;
  });
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) return i + 1;
  }
  return -1;
}

function getParentIdColIndex(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf('parentId') + 1;
}

function getCategoryFullPath(catName, allCategories) {
  return catName;
}

function getDescendantNames(catId, allCategories) {
  var names = [];
  var queue = [catId];
  while (queue.length > 0) {
    var current = queue.shift();
    for (var i = 0; i < allCategories.length; i++) {
      if (String(allCategories[i].parentId) === String(current)) {
        names.push(allCategories[i].name);
        queue.push(allCategories[i].id);
      }
    }
  }
  return names;
}

// ── GET handler (public — no auth required) ─────────────────

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';

  try {
    switch (action) {
      case 'getCategories': {
        const sheet = getSheet(CATEGORIES_SHEET);
        ensureCategoryParentIdColumn(sheet);
        return jsonResponse({ success: true, categories: sheetToObjects(sheet) });
      }

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
        const catSheet = getSheet(CATEGORIES_SHEET);
        ensureCategoryParentIdColumn(catSheet);
        const categories = sheetToObjects(catSheet);
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

      case 'getTranslations': {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName(TRANSLATIONS_SHEET);
        if (!sheet) {
          return jsonResponse({ success: true, translations: {} });
        }
        const data = sheet.getDataRange().getValues();
        if (data.length < 2) return jsonResponse({ success: true, translations: {} });
        const headers = data[0]; // key, el, en, ...
        var result = {};
        for (var langIdx = 1; langIdx < headers.length; langIdx++) {
          var lang = String(headers[langIdx]).toLowerCase().trim();
          if (!lang) continue;
          result[lang] = {};
          for (var r = 1; r < data.length; r++) {
            var key = String(data[r][0]).trim();
            if (key) result[lang][key] = String(data[r][langIdx] || '');
          }
        }
        return jsonResponse({ success: true, translations: result });
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

    const auth = requireEditor(body);
    if (!auth.ok) {
      return jsonResponse({ success: false, error: auth.error, message: auth.message });
    }

    switch (action) {

      // ── Categories ──────────────────────────────────────

      case 'addCategory': {
        const sheet = getSheet(CATEGORIES_SHEET);
        ensureCategoryParentIdColumn(sheet);
        const id = Utilities.getUuid();
        const parentId = body.parentId || '';
        sheet.appendRow([id, body.name, parentId]);
        return jsonResponse({ success: true, category: { id: id, name: body.name, parentId: parentId } });
      }

      case 'renameCategory': {
        const sheet = getSheet(CATEGORIES_SHEET);
        ensureCategoryParentIdColumn(sheet);
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
        ensureCategoryParentIdColumn(sheet);
        const row = findRowById(sheet, body.id);
        if (row === -1) return jsonResponse({ success: false, error: 'Category not found' });
        const catName = sheet.getRange(row, 2).getValue();
        const allCats = sheetToObjects(sheet);
        var descendantNames = getDescendantNames(body.id, allCats);
        var allNamesToDelete = [catName].concat(descendantNames);
        const itemsSheet = getSheet(ITEMS_SHEET);
        const itemData = itemsSheet.getDataRange().getValues();
        var itemCount = 0;
        for (var i = 1; i < itemData.length; i++) {
          if (allNamesToDelete.indexOf(itemData[i][2]) !== -1) itemCount++;
        }
        if (itemCount > 0 && !body.force) {
          return jsonResponse({ success: false, error: 'HAS_ITEMS', itemCount: itemCount, categoryName: catName });
        }
        if (body.force && body.deleteItems) {
          for (var i = itemData.length - 1; i >= 1; i--) {
            if (allNamesToDelete.indexOf(itemData[i][2]) !== -1) {
              itemsSheet.deleteRow(i + 1);
            }
          }
        }
        if (body.force && body.reassignTo) {
          for (var i = 1; i < itemData.length; i++) {
            if (allNamesToDelete.indexOf(itemData[i][2]) !== -1) {
              itemsSheet.getRange(i + 1, 3).setValue(body.reassignTo);
              itemsSheet.getRange(i + 1, 8).setValue(new Date().toISOString());
            }
          }
        }
        // Delete descendants first (reverse order to keep row indices valid)
        var descIds = [];
        var queue = [body.id];
        while (queue.length > 0) {
          var cur = queue.shift();
          for (var j = 0; j < allCats.length; j++) {
            if (String(allCats[j].parentId) === String(cur)) {
              descIds.push(allCats[j].id);
              queue.push(allCats[j].id);
            }
          }
        }
        var allIdsToDelete = descIds.concat([body.id]);
        // Delete from bottom up
        var catData = sheet.getDataRange().getValues();
        for (var i = catData.length - 1; i >= 1; i--) {
          if (allIdsToDelete.indexOf(catData[i][0]) !== -1) {
            sheet.deleteRow(i + 1);
          }
        }
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

      case 'seedTranslations': {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName(TRANSLATIONS_SHEET);
        if (!sheet) {
          sheet = ss.insertSheet(TRANSLATIONS_SHEET);
        }
        var translations = body.translations; // { key: { el: '...', en: '...' } }
        var keys = Object.keys(translations);
        var langs = ['el', 'en'];
        // Build data array
        var data = [['key'].concat(langs)];
        keys.forEach(function(key) {
          var row = [key];
          langs.forEach(function(lang) {
            row.push(translations[key][lang] || '');
          });
          data.push(row);
        });
        sheet.clear();
        sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
        return jsonResponse({ success: true });
      }

      case 'importAll': {
        if (body.categories) {
          const catSheet = getSheet(CATEGORIES_SHEET);
          catSheet.clear();
          catSheet.appendRow(['id', 'name', 'parentId']);
          body.categories.forEach(c => catSheet.appendRow([c.id, c.name, c.parentId || '']));
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
