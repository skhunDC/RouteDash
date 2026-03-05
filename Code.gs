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
  return HtmlService.createTemplate(
    '<h1>Missing index.html</h1><p>Add index.html to the Apps Script project.</p>'
  );
}

var DASHBOARD_LOG_SHEET_NAME = 'Route Dashboard Logs';
var DASHBOARD_FRAME_SHEET_NAME = 'Route Dashboard Frames';

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
  // Example using Open-Meteo. Replace latitude and longitude with your own.
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

function getFrames() {
  var props = PropertiesService.getScriptProperties();
  var frames = getFramesFromSheet();
  if (frames && frames.length) {
    return frames;
  }
  var data = props.getProperty('frames');
  return data ? JSON.parse(data) : [];
}

function saveFrames(frames) {
  var safeFrames = Array.isArray(frames) ? frames : [];
  PropertiesService.getScriptProperties().setProperty('frames', JSON.stringify(safeFrames));
  saveFramesToSheet(safeFrames);
  logDashboardUpdate('Saved frames', 'Count: ' + (frames ? frames.length : 0));
}

function getOrCreateFrameSheet() {
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty('frameSpreadsheetId');
  var spreadsheet = null;

  if (sheetId) {
    try {
      spreadsheet = SpreadsheetApp.openById(sheetId);
    } catch (e) {
      spreadsheet = null;
    }
  }

  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.create(DASHBOARD_FRAME_SHEET_NAME);
    props.setProperty('frameSpreadsheetId', spreadsheet.getId());
  }

  var sheet = spreadsheet.getSheetByName(DASHBOARD_FRAME_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(DASHBOARD_FRAME_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet
      .getRange(1, 1, 1, 7)
      .setValues([['Frame Id', 'Title', 'Rows', 'Columns', 'Sheet JSON', 'Dashboard JSON', 'Updated At']]);
  }

  return sheet;
}

function saveFramesToSheet(frames) {
  try {
    var sheet = getOrCreateFrameSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 7).clearContent();
    }
    if (!frames.length) return;

    var values = frames.map(function(frame) {
      var sheetState = frame && frame.sheet ? frame.sheet : {};
      var rows = Number(sheetState.rows) || 0;
      var cols = Number(sheetState.cols) || 0;
      return [
        frame.id || '',
        frame.title || 'Frame',
        rows,
        cols,
        JSON.stringify(sheetState),
        JSON.stringify(frame.dashboard || {}),
        new Date()
      ];
    });

    sheet.getRange(2, 1, values.length, 7).setValues(values);
  } catch (e) {
    // keep Script Properties as fallback when Sheet sync fails
  }
}

function getFramesFromSheet() {
  try {
    var sheet = getOrCreateFrameSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    return values
      .map(function(row) {
        var id = row[0] || '';
        var title = row[1] || 'Frame';
        var sheetJson = row[4] || '';
        var dashboardJson = row[5] || '';
        if (!sheetJson) return null;
        var sheetState = null;
        var dashboardState = {};
        try {
          sheetState = JSON.parse(sheetJson);
        } catch (e) {
          return null;
        }
        if (dashboardJson) {
          try {
            dashboardState = JSON.parse(dashboardJson);
          } catch (e) {
            dashboardState = {};
          }
        }
        return {
          id: id,
          title: title,
          sheet: sheetState,
          dashboard: dashboardState,
          x: 0,
          y: 0,
          width: 320,
          height: 360
        };
      })
      .filter(function(frame) {
        return frame !== null;
      });
  } catch (e) {
    return [];
  }
}

function getFloatingLayouts() {
  var props = PropertiesService.getScriptProperties();
  var data = props.getProperty('floatingLayouts');
  if (!data) return {};
  try {
    return JSON.parse(data);
  } catch (e) {
    return {};
  }
}

function saveFloatingLayouts(layouts) {
  var props = PropertiesService.getScriptProperties();
  if (!layouts || typeof layouts !== 'object') {
    props.deleteProperty('floatingLayouts');
    logDashboardUpdate('Cleared floating layouts');
    return;
  }
  props.setProperty('floatingLayouts', JSON.stringify(layouts));
  logDashboardUpdate('Saved floating layouts');
}

function getDriverOfTheWeek() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty('driverOfWeek') || '';
}

function saveDriverOfTheWeek(name) {
  PropertiesService.getScriptProperties().setProperty('driverOfWeek', name || '');
  logDashboardUpdate('Updated driver of the week', name || '');
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
      // fetch failed; use a random fallback without updating the timestamp
      quote = FALLBACK_QUOTES[Math.floor(Math.random() * FALLBACK_QUOTES.length)];
    }
  }

  return quote || FALLBACK_QUOTES[Math.floor(Math.random() * FALLBACK_QUOTES.length)];
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
  // return null to indicate failure so caller can decide how to handle
  return null;
}

function getHeaderTitle() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty('headerTitle') || 'Route Operations Dashboard';
}

function saveHeaderTitle(title) {
  PropertiesService.getScriptProperties().setProperty(
    'headerTitle',
    title || 'Route Operations Dashboard'
  );
  logDashboardUpdate('Updated header title', title || 'Route Operations Dashboard');
}

function getLogoImage() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty('logoImage') || '';
}

function saveLogoImage(data) {
  var props = PropertiesService.getScriptProperties();
  if (data) {
    props.setProperty('logoImage', data);
    logDashboardUpdate('Updated logo image');
  } else {
    props.deleteProperty('logoImage');
    logDashboardUpdate('Cleared logo image');
  }
}

function getDriverImages() {
  var props = PropertiesService.getScriptProperties();
  var ids = props.getProperty('driverImageIds');
  if (!ids) return [];
  ids = JSON.parse(ids);
  var images = [];
  var valid = [];
  ids.forEach(function(id) {
    try {
      var file = DriveApp.getFileById(id);
      var blob = file.getBlob();
      var base64 = Utilities.base64Encode(blob.getBytes());
      images.push('data:' + blob.getContentType() + ';base64,' + base64);
      valid.push(id);
    } catch (e) {
      // ignore missing file
    }
  });
  if (valid.length !== ids.length) {
    props.setProperty('driverImageIds', JSON.stringify(valid));
  }
  return images;
}

function addDriverImage(data) {
  if (!data) return;
  var m = data.match(/^data:(.+);base64,(.+)$/);
  if (!m) return;
  var props = PropertiesService.getScriptProperties();
  var ids = props.getProperty('driverImageIds');
  var arr = ids ? JSON.parse(ids) : [];
  var contentType = m[1];
  var bytes = Utilities.base64Decode(m[2]);
  var blob = Utilities.newBlob(bytes, contentType, 'driver-image');
  var file = DriveApp.createFile(blob);
  arr.push(file.getId());
  props.setProperty('driverImageIds', JSON.stringify(arr));
  logDashboardUpdate('Added driver image', 'Count: ' + arr.length);
}

function clearDriverImages() {
  var props = PropertiesService.getScriptProperties();
  var ids = props.getProperty('driverImageIds');
  if (!ids) return;
  JSON.parse(ids).forEach(function(id) {
    try {
      DriveApp.getFileById(id).setTrashed(true);
    } catch (e) {
      // ignore missing file
    }
  });
  props.deleteProperty('driverImageIds');
  logDashboardUpdate('Cleared driver images');
}

function updateDriverImages(list) {
  // Replace existing driver images with a new set provided as base64 data URLs
  clearDriverImages();
  if (!list || !list.length) return [];
  var ids = [];
  list.forEach(function(data) {
    var m = data.match(/^data:(.+);base64,(.+)$/);
    if (!m) return;
    var contentType = m[1];
    var bytes = Utilities.base64Decode(m[2]);
    var blob = Utilities.newBlob(bytes, contentType, 'driver-image');
    var file = DriveApp.createFile(blob);
    ids.push(file.getId());
  });
  PropertiesService.getScriptProperties().setProperty(
    'driverImageIds',
    JSON.stringify(ids)
  );
  logDashboardUpdate('Updated driver images', 'Count: ' + ids.length);
  return getDriverImages();
}

var APP_DATA_SPREADSHEET_NAME = 'Route Dashboard App Data';
var COMPETITION_CATEGORIES_SHEET = 'Competition_Categories';
var COMPETITION_ENTRIES_SHEET = 'Competition_Entries';
var APP_LINKS_SHEET = 'App_Links';
var DEFAULT_PATIO_LINK = 'https://docs.google.com/spreadsheets/d/1_wQmmrvmFD41CunrE6NUw1no1m4-MPghH4kDmBqZxps/edit?usp=sharing';

function getOrCreateAppDataSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty('appDataSpreadsheetId');
  var spreadsheet = null;
  if (sheetId) {
    try {
      spreadsheet = SpreadsheetApp.openById(sheetId);
    } catch (e) {
      spreadsheet = null;
    }
  }

  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.create(APP_DATA_SPREADSHEET_NAME);
    props.setProperty('appDataSpreadsheetId', spreadsheet.getId());
  }
  return spreadsheet;
}

function getOrCreateSheetWithHeader(sheetName, headers) {
  var spreadsheet = getOrCreateAppDataSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function getCompetitionCategories() {
  var sheet = getOrCreateSheetWithHeader(COMPETITION_CATEGORIES_SHEET, [
    'CategoryKey', 'CategoryName', 'Enabled', 'SortOrder', 'Goal', 'Notes', 'UpdatedAt'
  ]);
  var lastRow = sheet.getLastRow();
  var all = [];
  if (lastRow > 1) {
    all = sheet.getRange(2, 1, lastRow - 1, 7).getValues().map(function(row) {
      return {
        CategoryKey: row[0] || '',
        CategoryName: row[1] || '',
        Enabled: row[2] === true || String(row[2]).toUpperCase() === 'TRUE',
        SortOrder: Number(row[3]) || 0,
        Goal: row[4] || '',
        Notes: row[5] || '',
        UpdatedAt: row[6] || ''
      };
    }).filter(function(category) {
      return category.CategoryKey;
    });
  }

  if (!all.length) {
    all = [
      { CategoryKey: 'conversions', CategoryName: 'CSR Conversions', Enabled: true, SortOrder: 1, Goal: '', Notes: '', UpdatedAt: new Date() },
      { CategoryKey: 'patio_signups', CategoryName: 'Patio Signups', Enabled: true, SortOrder: 2, Goal: '', Notes: '', UpdatedAt: new Date() },
      { CategoryKey: 'alterations', CategoryName: 'Alterations', Enabled: true, SortOrder: 3, Goal: '', Notes: '', UpdatedAt: new Date() }
    ];
    saveCompetitionCategories(all);
  }

  all.sort(function(a, b) {
    return a.SortOrder - b.SortOrder;
  });
  return {
    allCategories: all,
    enabledCategories: all.filter(function(category) { return category.Enabled; })
  };
}

function saveCompetitionCategories(payload) {
  var categories = Array.isArray(payload) ? payload : [];
  var sheet = getOrCreateSheetWithHeader(COMPETITION_CATEGORIES_SHEET, [
    'CategoryKey', 'CategoryName', 'Enabled', 'SortOrder', 'Goal', 'Notes', 'UpdatedAt'
  ]);
  var now = new Date();
  var rows = categories.map(function(category, index) {
    var key = (category.CategoryKey || '').toString().trim();
    return [
      key,
      category.CategoryName || key,
      category.Enabled !== false,
      Number(category.SortOrder) || (index + 1),
      category.Goal || '',
      category.Notes || '',
      now
    ];
  }).filter(function(row) {
    return row[0];
  });
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 7).clearContent();
  }
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }
  return { success: true, count: rows.length, updatedAt: now };
}

function getCompetitionCategoryData(categoryKey, dateRange, store) {
  var key = (categoryKey || '').toString().trim();
  var sheet = getOrCreateSheetWithHeader(COMPETITION_ENTRIES_SHEET, [
    'Date', 'Store', 'CSR', 'CategoryKey', 'Value', 'Notes', 'UpdatedAt'
  ]);
  var dateToken = dateRange || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT', 'yyyy-MM-dd');
  var lastRow = sheet.getLastRow();
  var rows = [];
  if (lastRow > 1) {
    rows = sheet.getRange(2, 1, lastRow - 1, 7).getValues().filter(function(row) {
      if (key && row[3] !== key) return false;
      if (dateToken && row[0] !== dateToken) return false;
      if (store && row[1] !== store) return false;
      return true;
    }).map(function(row) {
      return {
        Date: row[0] || '',
        Store: row[1] || '',
        CSR: row[2] || '',
        CategoryKey: row[3] || '',
        Value: Number(row[4]) || 0,
        Notes: row[5] || '',
        UpdatedAt: row[6] || ''
      };
    });
  }
  rows.sort(function(a, b) {
    return b.Value - a.Value;
  });
  return {
    categoryKey: key,
    dateRange: dateToken,
    store: store || '',
    rows: rows,
    lastUpdated: rows.length ? rows[0].UpdatedAt : ''
  };
}

function saveCompetitionEntries(payload) {
  var entries = Array.isArray(payload) ? payload : [];
  var sheet = getOrCreateSheetWithHeader(COMPETITION_ENTRIES_SHEET, [
    'Date', 'Store', 'CSR', 'CategoryKey', 'Value', 'Notes', 'UpdatedAt'
  ]);
  var targetDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT', 'yyyy-MM-dd');
  var storeFilter = '';
  if (entries.length) {
    targetDate = entries[0].Date || targetDate;
    storeFilter = entries[0].Store || '';
  }

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var existing = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    var removeIndexes = [];
    existing.forEach(function(row, idx) {
      if (row[0] === targetDate && row[1] === storeFilter) {
        removeIndexes.push(idx + 2);
      }
    });
    for (var i = removeIndexes.length - 1; i >= 0; i--) {
      sheet.deleteRow(removeIndexes[i]);
    }
  }

  var now = new Date();
  var rows = entries.map(function(entry) {
    return [
      entry.Date || targetDate,
      entry.Store || '',
      entry.CSR || '',
      entry.CategoryKey || '',
      Number(entry.Value) || 0,
      entry.Notes || '',
      now
    ];
  }).filter(function(row) {
    return row[2] && row[3];
  });
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
  }
  return { success: true, count: rows.length, updatedAt: now };
}

function getAppLinks() {
  var sheet = getOrCreateSheetWithHeader(APP_LINKS_SHEET, ['Key', 'Label', 'Url', 'UpdatedAt']);
  var lastRow = sheet.getLastRow();
  var links = {};
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).getValues().forEach(function(row) {
      if (!row[0]) return;
      links[row[0]] = {
        label: row[1] || row[0],
        url: row[2] || '',
        updatedAt: row[3] || ''
      };
    });
  }
  if (!links.PATIO_CUSHION_SIGNUP) {
    saveAppLinks([{ Key: 'PATIO_CUSHION_SIGNUP', Label: 'Patio Cushion Signup Sheet', Url: DEFAULT_PATIO_LINK }]);
    links.PATIO_CUSHION_SIGNUP = { label: 'Patio Cushion Signup Sheet', url: DEFAULT_PATIO_LINK, updatedAt: new Date() };
  }
  return links;
}

function saveAppLinks(payload) {
  var incoming = Array.isArray(payload) ? payload : [];
  var sheet = getOrCreateSheetWithHeader(APP_LINKS_SHEET, ['Key', 'Label', 'Url', 'UpdatedAt']);
  var now = new Date();
  var byKey = {};
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).getValues().forEach(function(row) {
      if (!row[0]) return;
      byKey[row[0]] = { Key: row[0], Label: row[1], Url: row[2], UpdatedAt: row[3] };
    });
  }
  incoming.forEach(function(item) {
    if (!item || !item.Key) return;
    byKey[item.Key] = {
      Key: item.Key,
      Label: item.Label || item.Key,
      Url: item.Url || '',
      UpdatedAt: now
    };
  });
  var rows = Object.keys(byKey).sort().map(function(key) {
    var item = byKey[key];
    return [item.Key, item.Label, item.Url, item.UpdatedAt || now];
  });
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
  return { success: true, count: rows.length, updatedAt: now };
}

function getDashboardData() {
  var categoriesPayload = getCompetitionCategories();
  var enabled = categoriesPayload.enabledCategories;
  var summary = enabled.map(function(category) {
    var data = getCompetitionCategoryData(category.CategoryKey);
    var leader = data.rows.length ? data.rows[0] : null;
    return {
      categoryKey: category.CategoryKey,
      categoryName: category.CategoryName,
      leader: leader ? leader.CSR : '',
      leaderValue: leader ? leader.Value : 0,
      goal: category.Goal || ''
    };
  });
  var links = getAppLinks();
  return {
    competitionsSummary: summary,
    patioCushionSignupLink: links.PATIO_CUSHION_SIGNUP || { label: 'Patio Cushion Signup Sheet', url: DEFAULT_PATIO_LINK }
  };
}
