function doGet() {
  var spreadsheetId = '1d4XuVJPSy579inDl082_Qr84CHU5rFpcnx0kSwiblg4';
  var sheetName = 'シート1';
  var settingsSheetName = 'settings';

  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Sheet not found.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) {
    return ContentService
      .createTextOutput(JSON.stringify({ rows: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var range = sheet.getRange(1, 1, lastRow, lastColumn);
  var values = range.getDisplayValues();
  var backgrounds = range.getBackgrounds();
  var fontColors = range.getFontColors();
  var fontFamilies = range.getFontFamilies();
  var fontSizes = range.getFontSizes();
  var fontWeights = range.getFontWeights();
  var fontStyles = range.getFontStyles();
  var fontLines = range.getFontLines();
  var richTextValues = range.getRichTextValues();
  var headers = values[0];

  var normalizedHeaders = headers.map(function(header) {
    return String(header).trim().toLowerCase().replace(/[\s-]+/g, '_');
  });

  var stepColumnIndex = normalizedHeaders.indexOf('step');
  var bodyColumnIndex = normalizedHeaders.indexOf('body');
  var rows = [];

  for (var rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    var rowValues = values[rowIndex];
    var row = {};
    var hasAnyValue = false;

    for (var columnIndex = 0; columnIndex < headers.length; columnIndex += 1) {
      var key = headers[columnIndex];
      if (!key) continue;

      var value = rowValues[columnIndex];
      row[key] = value;
      if (value !== '') {
        hasAnyValue = true;
      }
    }

    if (!hasAnyValue) {
      continue;
    }

    if (stepColumnIndex >= 0) {
      row.stepStyle = {
        backgroundColor: backgrounds[rowIndex][stepColumnIndex],
        textColor: fontColors[rowIndex][stepColumnIndex],
        fontFamily: fontFamilies[rowIndex][stepColumnIndex],
        fontSize: fontSizes[rowIndex][stepColumnIndex],
        bold: String(fontWeights[rowIndex][stepColumnIndex]).toLowerCase() === 'bold',
        italic: String(fontStyles[rowIndex][stepColumnIndex]).toLowerCase() === 'italic',
        underline: String(fontLines[rowIndex][stepColumnIndex]).toLowerCase() === 'underline',
        strikethrough:
          String(fontLines[rowIndex][stepColumnIndex]).toLowerCase() === 'line-through',
      };
      row.stepRuns = extractRuns(richTextValues[rowIndex][stepColumnIndex]);
    }

    if (bodyColumnIndex >= 0) {
      row.bodyStyle = {
        backgroundColor: backgrounds[rowIndex][bodyColumnIndex],
        textColor: fontColors[rowIndex][bodyColumnIndex],
        fontFamily: fontFamilies[rowIndex][bodyColumnIndex],
        fontSize: fontSizes[rowIndex][bodyColumnIndex],
        bold: String(fontWeights[rowIndex][bodyColumnIndex]).toLowerCase() === 'bold',
        italic: String(fontStyles[rowIndex][bodyColumnIndex]).toLowerCase() === 'italic',
        underline: String(fontLines[rowIndex][bodyColumnIndex]).toLowerCase() === 'underline',
        strikethrough:
          String(fontLines[rowIndex][bodyColumnIndex]).toLowerCase() === 'line-through',
      };
      row.bodyRuns = extractRuns(richTextValues[rowIndex][bodyColumnIndex]);
    }

    rows.push(row);
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      rows: rows,
      settings: getSettings_(spreadsheet, settingsSheetName),
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function extractRuns(richTextValue) {
  if (!richTextValue) {
    return [];
  }

  var runs = richTextValue.getRuns();
  if (!runs || runs.length === 0) {
    return [];
  }

  return runs
    .map(function(run) {
      var text = run.getText();
      if (!text) {
        return null;
      }

      var style = run.getTextStyle();
      return {
        text: text,
        textColor: style ? style.getForegroundColor() : '',
        fontFamily: style ? style.getFontFamily() : '',
        fontSize: style ? style.getFontSize() : '',
        bold: style ? style.isBold() : false,
        italic: style ? style.isItalic() : false,
        underline: style ? style.isUnderline() : false,
        strikethrough: style ? style.isStrikethrough() : false,
      };
    })
    .filter(function(run) {
      return run !== null;
    });
}

function getSettings_(spreadsheet, settingsSheetName) {
  var sheet = spreadsheet.getSheetByName(settingsSheetName);
  if (!sheet) {
    return {};
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    return {};
  }

  var values = sheet.getRange(1, 1, lastRow, 2).getDisplayValues();
  var settings = {};

  values.forEach(function(row) {
    var key = String(row[0] || '').trim();
    var value = String(row[1] || '').trim();
    if (!key) {
      return;
    }
    settings[key] = value;
  });

  return settings;
}
