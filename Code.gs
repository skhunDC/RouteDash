function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('Route Operations Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

var DASHBOARD_LOG_SHEET_NAME = 'Route Dashboard Logs';

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
  var data = props.getProperty('frames');
  return data ? JSON.parse(data) : [];
}

function saveFrames(frames) {
  PropertiesService.getScriptProperties().setProperty('frames', JSON.stringify(frames));
  logDashboardUpdate('Saved frames', 'Count: ' + (frames ? frames.length : 0));
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
