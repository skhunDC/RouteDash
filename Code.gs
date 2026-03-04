function doGet() {
  var template = getDashboardTemplate();
  return template.evaluate()
    .setTitle('Route Operations Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardTemplate() {
  var names = ['index', 'Index'];
  for (var i = 0; i < names.length; i++) {
    try {
      return HtmlService.createTemplateFromFile(names[i]);
    } catch (e) {
      // try the next candidate
    }
  }
  return HtmlService.createTemplate('<h1>Missing index.html</h1><p>Add index.html to the Apps Script project.</p>');
}

var DASHBOARD_LOG_SHEET_NAME = 'Route Dashboard Logs';
var DATA_SPREADSHEET_PROP = 'dataSpreadsheetId';
var META_SHEET_NAME = 'META';
var IMAGES_SHEET_NAME = 'IMAGES';
var RESERVED_SHEETS = [META_SHEET_NAME, IMAGES_SHEET_NAME];

function getOrCreateLogSheet() {
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty('logSpreadsheetId');
  var spreadsheet = null;

  if (sheetId) {
    try {
      spreadsheet = SpreadsheetApp.openById(sheetId);
    } catch (e) {
      spreadsheet = null;
    }
  }

  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.create(DASHBOARD_LOG_SHEET_NAME);
    props.setProperty('logSpreadsheetId', spreadsheet.getId());
  }

  var sheet = spreadsheet.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'User', 'Action', 'Details']]);
  }
  return sheet;
}

function logDashboardUpdate(action, details) {
  if (!action) return;
  try {
    var sheet = getOrCreateLogSheet();
    var userEmail = Session.getActiveUser().getEmail() || 'Unknown';
    sheet.appendRow([new Date(), userEmail, action, details || '']);
  } catch (e) {
    // avoid blocking dashboard updates if logging fails
  }
}

function weatherCodeToText(code) {
  var mapping = {
    0: 'Clear sky',
    1: 'Mainly clear',
    2: 'Partly cloudy',
    3: 'Overcast',
    45: 'Fog',
    48: 'Depositing rime fog',
    51: 'Light drizzle',
    53: 'Moderate drizzle',
    55: 'Dense drizzle',
    56: 'Light freezing drizzle',
    57: 'Dense freezing drizzle',
    61: 'Slight rain',
    63: 'Moderate rain',
    65: 'Heavy rain',
    66: 'Light freezing rain',
    67: 'Heavy freezing rain',
    71: 'Slight snow fall',
    73: 'Moderate snow fall',
    75: 'Heavy snow fall',
    77: 'Snow grains',
    80: 'Slight rain showers',
    81: 'Moderate rain showers',
    82: 'Violent rain showers',
    85: 'Slight snow showers',
    86: 'Heavy snow showers',
    95: 'Thunderstorm',
    96: 'Thunderstorm with slight hail',
    99: 'Thunderstorm with heavy hail'
  };
  return mapping[code] || 'Unknown';
}

function getWeather() {
  var url = 'https://api.open-meteo.com/v1/forecast?latitude=40.7128&longitude=-74.0060&current_weather=true&temperature_unit=fahrenheit';
  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());
  if (data && data.current_weather) {
    return {
      temperature: data.current_weather.temperature,
      windspeed: data.current_weather.windspeed,
      condition: weatherCodeToText(data.current_weather.weathercode)
    };
  }
  return null;
}

var FALLBACK_QUOTES = [
  'The only way to do great work is to love what you do. — Steve Jobs',
  'Success is not final, failure is not fatal: it is the courage to continue that counts. — Winston Churchill',
  'Life is what happens when you\'re busy making other plans. — John Lennon',
  'You miss 100% of the shots you don\'t take. — Wayne Gretzky'
];

function fetchQuote() {
  var url = 'https://zenquotes.io/api/random';
  try {
    var response = UrlFetchApp.fetch(url);
    var data = JSON.parse(response.getContentText());
    if (Array.isArray(data) && data.length > 0 && data[0].q && data[0].a) {
      return data[0].q + ' — ' + data[0].a;
    }
  } catch (e) {
    // ignore errors and fall through
  }
  return null;
}

function getRandomQuote() {
  var props = PropertiesService.getScriptProperties();
  var quote = props.getProperty('dailyQuote');
  var timestamp = props.getProperty('dailyQuoteDate');
  var now = new Date();
  var needNew = true;

  if (quote && timestamp) {
    var last = new Date(timestamp);
    if (now.getTime() - last.getTime() < 24 * 60 * 60 * 1000) {
      needNew = false;
    }
  }

  if (needNew) {
    var newQuote = fetchQuote();
    if (newQuote) {
      props.setProperty('dailyQuote', newQuote);
      props.setProperty('dailyQuoteDate', now.toISOString());
      quote = newQuote;
    } else {
      quote = FALLBACK_QUOTES[Math.floor(Math.random() * FALLBACK_QUOTES.length)];
    }
  }

  return quote || FALLBACK_QUOTES[Math.floor(Math.random() * FALLBACK_QUOTES.length)];
}

function getOrCreateDataSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var spreadsheetId = props.getProperty(DATA_SPREADSHEET_PROP);
  var spreadsheet;

  if (spreadsheetId) {
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      spreadsheet = null;
    }
  }

  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.create('RouteDash Data Store');
    props.setProperty(DATA_SPREADSHEET_PROP, spreadsheet.getId());
  }

  ensureCoreSheets(spreadsheet);
  return spreadsheet.getId();
}

function ensureCoreSheets(spreadsheet) {
  var meta = spreadsheet.getSheetByName(META_SHEET_NAME);
  if (!meta) {
    meta = spreadsheet.insertSheet(META_SHEET_NAME);
  }
  if (meta.getLastRow() === 0) {
    meta.getRange(1, 1, 1, 2).setValues([['Key', 'JsonValue']]);
  }

  var images = spreadsheet.getSheetByName(IMAGES_SHEET_NAME);
  if (!images) {
    images = spreadsheet.insertSheet(IMAGES_SHEET_NAME);
  }
  if (images.getLastRow() === 0) {
    images.getRange(1, 1, 1, 6).setValues([['imageId', 'fileId', 'name', 'mimeType', 'createdAt', 'createdBy']]);
  }

  if (!spreadsheet.getSheetByName('Sheet1')) {
    spreadsheet.insertSheet('Sheet1');
  }

  var layout = readMetaValue(spreadsheet, 'dashboardLayout');
  if (!layout) {
    writeMetaValue(spreadsheet, 'dashboardLayout', defaultDashboardLayout());
  }

  var config = readMetaValue(spreadsheet, 'appConfig');
  if (!config) {
    writeMetaValue(spreadsheet, 'appConfig', defaultConfig());
  }
}

function defaultDashboardLayout() {
  return {
    mode: 'display',
    tiles: [
      { id: 'tile-weather', type: 'weather', title: 'Weather', size: 'small' },
      { id: 'tile-quote', type: 'quote', title: 'Daily Quote', size: 'wide' },
      { id: 'tile-image', type: 'image', title: 'Driver Spotlight', size: 'wide' },
      { id: 'tile-chart', type: 'chart', title: 'Sheet1 Overview', size: 'big', sheetName: 'Sheet1', range: 'A1:D12', chartType: 'bar' },
      { id: 'tile-sheet', type: 'sheet', title: 'Sheet1 Editor', size: 'big', sheetName: 'Sheet1', range: 'A1:F12' }
    ]
  };
}

function defaultConfig() {
  return {
    headerTitle: 'Route Operations Dashboard',
    driverOfWeek: '',
    rotationSeconds: 12
  };
}

function getDataSpreadsheet() {
  var spreadsheetId = getOrCreateDataSpreadsheet();
  return SpreadsheetApp.openById(spreadsheetId);
}

function readMetaValue(spreadsheet, key) {
  var metaSheet = spreadsheet.getSheetByName(META_SHEET_NAME);
  var lastRow = metaSheet.getLastRow();
  if (lastRow < 2) return null;
  var values = metaSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === key) {
      try {
        return JSON.parse(values[i][1]);
      } catch (e) {
        return null;
      }
    }
  }
  return null;
}

function writeMetaValue(spreadsheet, key, value) {
  var metaSheet = spreadsheet.getSheetByName(META_SHEET_NAME);
  var serialized = JSON.stringify(value);
  var lastRow = metaSheet.getLastRow();
  var foundRow = -1;

  if (lastRow >= 2) {
    var keys = metaSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (keys[i][0] === key) {
        foundRow = i + 2;
        break;
      }
    }
  }

  if (foundRow === -1) {
    metaSheet.appendRow([key, serialized]);
  } else {
    metaSheet.getRange(foundRow, 2).setValue(serialized);
  }
}

function getDataSheetNames(spreadsheet) {
  var sheets = spreadsheet.getSheets();
  var names = [];
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (RESERVED_SHEETS.indexOf(name) === -1) {
      names.push(name);
    }
  }
  return names;
}

function getAppState() {
  var spreadsheet = getDataSpreadsheet();
  var config = readMetaValue(spreadsheet, 'appConfig') || defaultConfig();
  var layout = readMetaValue(spreadsheet, 'dashboardLayout') || defaultDashboardLayout();
  var tabs = getDataSheetNames(spreadsheet);

  return {
    spreadsheetId: spreadsheet.getId(),
    config: config,
    layout: layout,
    sheetTabs: tabs,
    imagesIndex: getImagesIndex(),
    weather: getWeather(),
    quote: getRandomQuote(),
    timestamp: new Date().toISOString()
  };
}

function saveAppLayout(layout) {
  var spreadsheet = getDataSpreadsheet();
  writeMetaValue(spreadsheet, 'dashboardLayout', layout || defaultDashboardLayout());
  logDashboardUpdate('Saved app layout', 'Tiles: ' + ((layout && layout.tiles && layout.tiles.length) || 0));
}

function saveAppConfig(configPatch) {
  var spreadsheet = getDataSpreadsheet();
  var config = readMetaValue(spreadsheet, 'appConfig') || defaultConfig();
  var keys = Object.keys(configPatch || {});
  for (var i = 0; i < keys.length; i++) {
    config[keys[i]] = configPatch[keys[i]];
  }
  writeMetaValue(spreadsheet, 'appConfig', config);
  logDashboardUpdate('Updated app config', JSON.stringify(configPatch || {}));
  return config;
}

function getSheetRange(sheetName, a1Range, options) {
  var spreadsheet = getDataSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }

  var range = sheet.getRange(a1Range || 'A1:Z30');
  var includeFormulas = options && options.includeFormulas;
  var payload = {
    sheetName: sheetName,
    range: range.getA1Notation(),
    displayValues: range.getDisplayValues()
  };

  if (includeFormulas) {
    payload.formulas = range.getFormulas();
    payload.values = range.getValues();
  }

  return payload;
}

function applySheetEdits(edits) {
  if (!edits || !edits.length) {
    return { updated: 0 };
  }

  var spreadsheet = getDataSpreadsheet();
  var updated = 0;

  for (var i = 0; i < edits.length; i++) {
    var edit = edits[i];
    if (!edit || !edit.sheetName || !edit.row || !edit.col) {
      continue;
    }

    var sheet = spreadsheet.getSheetByName(edit.sheetName);
    if (!sheet) {
      continue;
    }

    var cell = sheet.getRange(Number(edit.row), Number(edit.col));
    var inputText = edit.inputText;
    if (typeof inputText !== 'string') {
      inputText = inputText === null || inputText === undefined ? '' : String(inputText);
    }

    if (inputText && inputText.charAt(0) === '=') {
      cell.setFormula(inputText);
    } else {
      cell.setValue(inputText);
    }
    updated += 1;
  }

  logDashboardUpdate('Applied sheet edits', 'Count: ' + updated);
  return { updated: updated };
}

function createSheetTab(name) {
  var spreadsheet = getDataSpreadsheet();
  var safeName = (name || '').trim();
  if (!safeName) throw new Error('Sheet name is required.');
  if (spreadsheet.getSheetByName(safeName)) throw new Error('Sheet already exists: ' + safeName);
  if (RESERVED_SHEETS.indexOf(safeName) !== -1) throw new Error('Name is reserved.');
  spreadsheet.insertSheet(safeName);
  logDashboardUpdate('Created sheet tab', safeName);
  return getDataSheetNames(spreadsheet);
}

function renameSheetTab(oldName, newName) {
  var spreadsheet = getDataSpreadsheet();
  var oldSafe = (oldName || '').trim();
  var newSafe = (newName || '').trim();
  if (!oldSafe || !newSafe) throw new Error('Both old and new names are required.');
  if (RESERVED_SHEETS.indexOf(oldSafe) !== -1) throw new Error('Cannot rename reserved tab: ' + oldSafe);
  if (RESERVED_SHEETS.indexOf(newSafe) !== -1) throw new Error('Target name is reserved.');
  var sheet = spreadsheet.getSheetByName(oldSafe);
  if (!sheet) throw new Error('Sheet not found: ' + oldSafe);
  if (spreadsheet.getSheetByName(newSafe)) throw new Error('Sheet already exists: ' + newSafe);
  sheet.setName(newSafe);
  logDashboardUpdate('Renamed sheet tab', oldSafe + ' => ' + newSafe);
  return getDataSheetNames(spreadsheet);
}

function deleteSheetTab(name) {
  var spreadsheet = getDataSpreadsheet();
  var safeName = (name || '').trim();
  if (!safeName) throw new Error('Sheet name is required.');
  if (RESERVED_SHEETS.indexOf(safeName) !== -1) throw new Error('Cannot delete reserved tab: ' + safeName);
  var sheet = spreadsheet.getSheetByName(safeName);
  if (!sheet) throw new Error('Sheet not found: ' + safeName);

  var dataTabs = getDataSheetNames(spreadsheet);
  if (dataTabs.length <= 1) {
    throw new Error('At least one data tab must remain.');
  }

  spreadsheet.deleteSheet(sheet);
  logDashboardUpdate('Deleted sheet tab', safeName);
  return getDataSheetNames(spreadsheet);
}

function getImagesIndex() {
  var spreadsheet = getDataSpreadsheet();
  var sheet = spreadsheet.getSheetByName(IMAGES_SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var list = [];
  for (var i = 0; i < rows.length; i++) {
    if (!rows[i][0] || !rows[i][1]) continue;
    list.push({
      imageId: rows[i][0],
      fileId: rows[i][1],
      name: rows[i][2] || 'Driver image',
      mimeType: rows[i][3] || 'image/jpeg',
      createdAt: rows[i][4],
      createdBy: rows[i][5]
    });
  }
  return list;
}

function addImages(base64List) {
  if (!base64List || !base64List.length) return getImagesIndex();
  var spreadsheet = getDataSpreadsheet();
  var sheet = spreadsheet.getSheetByName(IMAGES_SHEET_NAME);
  var user = Session.getActiveUser().getEmail() || 'Unknown';

  for (var i = 0; i < base64List.length; i++) {
    var entry = base64List[i];
    var m = String(entry || '').match(/^data:(.+);base64,(.+)$/);
    if (!m) continue;

    var mimeType = m[1];
    var bytes = Utilities.base64Decode(m[2]);
    var extension = mimeType.split('/')[1] || 'img';
    var name = 'route-image-' + new Date().getTime() + '-' + i + '.' + extension;
    var blob = Utilities.newBlob(bytes, mimeType, name);
    var file = DriveApp.createFile(blob);
    var imageId = Utilities.getUuid();
    sheet.appendRow([imageId, file.getId(), name, mimeType, new Date().toISOString(), user]);
  }

  logDashboardUpdate('Added images', 'Count: ' + base64List.length);
  return getImagesIndex();
}

function deleteImage(imageId) {
  if (!imageId) return getImagesIndex();
  var spreadsheet = getDataSpreadsheet();
  var sheet = spreadsheet.getSheetByName(IMAGES_SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] === imageId) {
      try {
        DriveApp.getFileById(rows[i][1]).setTrashed(true);
      } catch (e) {
        // ignore missing file
      }
      sheet.deleteRow(i + 2);
      logDashboardUpdate('Deleted image', imageId);
      break;
    }
  }
  return getImagesIndex();
}

function getImageById(imageId) {
  if (!imageId) return '';
  var images = getImagesIndex();
  for (var i = 0; i < images.length; i++) {
    if (images[i].imageId === imageId) {
      var file = DriveApp.getFileById(images[i].fileId);
      var blob = file.getBlob();
      var base64 = Utilities.base64Encode(blob.getBytes());
      return {
        imageId: imageId,
        name: images[i].name,
        dataUrl: 'data:' + blob.getContentType() + ';base64,' + base64
      };
    }
  }
  return '';
}

function migrateLegacyData() {
  var props = PropertiesService.getScriptProperties();
  var spreadsheet = getDataSpreadsheet();

  var layout = readMetaValue(spreadsheet, 'dashboardLayout') || defaultDashboardLayout();
  var config = readMetaValue(spreadsheet, 'appConfig') || defaultConfig();

  var legacyFramesRaw = props.getProperty('frames');
  if (legacyFramesRaw) {
    try {
      var legacyFrames = JSON.parse(legacyFramesRaw);
      if (legacyFrames && legacyFrames.length) {
        layout.tiles = [];
        for (var i = 0; i < legacyFrames.length; i++) {
          var frame = legacyFrames[i] || {};
          layout.tiles.push({
            id: frame.id || ('legacy-' + i),
            type: frame.dashboardHidden ? 'sheet' : 'chart',
            title: frame.name || ('Legacy Frame ' + (i + 1)),
            size: 'big',
            sheetName: frame.activeSheet || 'Sheet1',
            range: 'A1:F12',
            chartType: frame.chartType || 'bar'
          });
        }
      }
    } catch (e) {
      // ignore malformed legacy data
    }
  }

  var headerTitle = props.getProperty('headerTitle');
  if (headerTitle) config.headerTitle = headerTitle;
  var driverOfWeek = props.getProperty('driverOfWeek');
  if (driverOfWeek) config.driverOfWeek = driverOfWeek;

  var legacyImageIdsRaw = props.getProperty('driverImageIds');
  if (legacyImageIdsRaw) {
    try {
      var legacyImageIds = JSON.parse(legacyImageIdsRaw);
      var current = getImagesIndex();
      var existingByFileId = {};
      for (var j = 0; j < current.length; j++) {
        existingByFileId[current[j].fileId] = true;
      }

      var imageSheet = spreadsheet.getSheetByName(IMAGES_SHEET_NAME);
      for (var k = 0; k < legacyImageIds.length; k++) {
        var legacyId = legacyImageIds[k];
        if (!legacyId || existingByFileId[legacyId]) continue;
        try {
          var legacyFile = DriveApp.getFileById(legacyId);
          imageSheet.appendRow([
            Utilities.getUuid(),
            legacyId,
            legacyFile.getName() || ('Legacy image ' + (k + 1)),
            legacyFile.getMimeType(),
            new Date().toISOString(),
            'Legacy migration'
          ]);
        } catch (e2) {
          // ignore missing file
        }
      }
    } catch (e3) {
      // ignore malformed list
    }
  }

  writeMetaValue(spreadsheet, 'dashboardLayout', layout);
  writeMetaValue(spreadsheet, 'appConfig', config);
  logDashboardUpdate('Migrated legacy data', 'Migration completed to Sheets-backed storage');

  return {
    migrated: true,
    spreadsheetId: spreadsheet.getId(),
    tiles: layout.tiles.length,
    images: getImagesIndex().length
  };
}

function getHeaderTitle() {
  var spreadsheet = getDataSpreadsheet();
  var config = readMetaValue(spreadsheet, 'appConfig') || defaultConfig();
  return config.headerTitle || 'Route Operations Dashboard';
}

function saveHeaderTitle(title) {
  return saveAppConfig({ headerTitle: title || 'Route Operations Dashboard' });
}

function getDriverOfTheWeek() {
  var spreadsheet = getDataSpreadsheet();
  var config = readMetaValue(spreadsheet, 'appConfig') || defaultConfig();
  return config.driverOfWeek || '';
}

function saveDriverOfTheWeek(name) {
  return saveAppConfig({ driverOfWeek: name || '' });
}

// Legacy wrappers for compatibility with older client calls.
function getDriverImages() {
  var index = getImagesIndex();
  var images = [];
  for (var i = 0; i < index.length; i++) {
    var image = getImageById(index[i].imageId);
    if (image && image.dataUrl) images.push(image.dataUrl);
  }
  return images;
}

function addDriverImage(data) {
  return addImages([data]);
}

function clearDriverImages() {
  var all = getImagesIndex();
  for (var i = 0; i < all.length; i++) {
    deleteImage(all[i].imageId);
  }
}

function updateDriverImages(list) {
  clearDriverImages();
  addImages(list || []);
  return getDriverImages();
}
