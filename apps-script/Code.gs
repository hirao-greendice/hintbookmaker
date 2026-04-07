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
      row.stepStyle = extractCellStyle_({
        backgroundColor: backgrounds[rowIndex][stepColumnIndex],
        textColor: fontColors[rowIndex][stepColumnIndex],
        fontFamily: fontFamilies[rowIndex][stepColumnIndex],
        fontSize: fontSizes[rowIndex][stepColumnIndex],
        fontWeight: fontWeights[rowIndex][stepColumnIndex],
        fontStyle: fontStyles[rowIndex][stepColumnIndex],
        fontLine: fontLines[rowIndex][stepColumnIndex],
      });
      row.stepRuns = extractRuns(
        richTextValues[rowIndex][stepColumnIndex],
        row.stepStyle
      );
    }

    if (bodyColumnIndex >= 0) {
      row.bodyStyle = extractCellStyle_({
        backgroundColor: backgrounds[rowIndex][bodyColumnIndex],
        textColor: fontColors[rowIndex][bodyColumnIndex],
        fontFamily: fontFamilies[rowIndex][bodyColumnIndex],
        fontSize: fontSizes[rowIndex][bodyColumnIndex],
        fontWeight: fontWeights[rowIndex][bodyColumnIndex],
        fontStyle: fontStyles[rowIndex][bodyColumnIndex],
        fontLine: fontLines[rowIndex][bodyColumnIndex],
      });
      row.bodyRuns = extractRuns(
        richTextValues[rowIndex][bodyColumnIndex],
        row.bodyStyle
      );
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

function extractCellStyle_(styleSeed) {
  var fontLine = String(styleSeed.fontLine || '').toLowerCase();
  return {
    backgroundColor: styleSeed.backgroundColor || '',
    textColor: styleSeed.textColor || '',
    fontFamily: styleSeed.fontFamily || '',
    fontSize: styleSeed.fontSize || '',
    bold: String(styleSeed.fontWeight || '').toLowerCase() === 'bold',
    italic: String(styleSeed.fontStyle || '').toLowerCase() === 'italic',
    underline: fontLine === 'underline',
    strikethrough: fontLine === 'line-through',
  };
}

function normalizeRunValue_(value) {
  return value === null || value === undefined ? '' : value;
}

function extractRuns(richTextValue, cellStyle) {
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
      var textColor = style ? style.getForegroundColor() : '';
      var fontFamily = style ? style.getFontFamily() : '';
      var fontSize = style ? style.getFontSize() : '';
      var bold = style ? style.isBold() : false;
      var italic = style ? style.isItalic() : false;
      var underline = style ? style.isUnderline() : false;
      var strikethrough = style ? style.isStrikethrough() : false;

      return {
        text: text,
        textColor:
          normalizeRunValue_(textColor) === normalizeRunValue_(cellStyle && cellStyle.textColor)
            ? ''
            : textColor,
        fontFamily:
          normalizeRunValue_(fontFamily) === normalizeRunValue_(cellStyle && cellStyle.fontFamily)
            ? ''
            : fontFamily,
        fontSize:
          normalizeRunValue_(fontSize) === normalizeRunValue_(cellStyle && cellStyle.fontSize)
            ? ''
            : fontSize,
        bold:
          normalizeRunValue_(bold) === normalizeRunValue_(cellStyle && cellStyle.bold)
            ? null
            : bold,
        italic:
          normalizeRunValue_(italic) === normalizeRunValue_(cellStyle && cellStyle.italic)
            ? null
            : italic,
        underline:
          normalizeRunValue_(underline) === normalizeRunValue_(cellStyle && cellStyle.underline)
            ? null
            : underline,
        strikethrough:
          normalizeRunValue_(strikethrough) ===
          normalizeRunValue_(cellStyle && cellStyle.strikethrough)
            ? null
            : strikethrough,
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
