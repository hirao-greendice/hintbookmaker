function doGet() {
  var spreadsheetId = '1d4XuVJPSy579inDl082_Qr84CHU5rFpcnx0kSwiblg4';
  var sheetName = 'シート1';
  var settingsSheetName = 'settings';
  var sideSheetName = 'side';

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
      .createTextOutput(JSON.stringify({
        rows: [],
        sideDefinitions: getSideDefinitions_(spreadsheet, sideSheetName),
      }))
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
    return normalizeHeader_(header);
  });

  var stepColumnIndex = normalizedHeaders.indexOf('step');
  var bodyColumnIndex = normalizedHeaders.indexOf('body');
  var imageColumnIndexes = normalizedHeaders
    .map(function(header, index) {
      return isImageHeader_(header) ? index : -1;
    })
    .filter(function(index) {
      return index >= 0;
    });
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

    imageColumnIndexes.forEach(function(columnIndex) {
      var imageLink = extractLinkUrl_(richTextValues[rowIndex][columnIndex]);
      if (imageLink) {
        row[headers[columnIndex]] = imageLink;
      }
    });

    rows.push(row);
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      rows: rows,
      settings: getSettings_(spreadsheet, settingsSheetName),
      sideDefinitions: getSideDefinitions_(spreadsheet, sideSheetName),
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

function normalizeHeader_(value) {
  return String(value || '').trim().toLowerCase().replace(/[\s-]+/g, '_');
}

function isImageHeader_(header) {
  return /^image(?:_\d+|\d+)?$/.test(String(header || '').trim().toLowerCase());
}

function extractLinkUrl_(richTextValue) {
  if (!richTextValue) {
    return '';
  }

  var directLink = richTextValue.getLinkUrl();
  if (directLink) {
    return directLink;
  }

  var runs = richTextValue.getRuns();
  if (!runs || runs.length === 0) {
    return '';
  }

  for (var index = 0; index < runs.length; index += 1) {
    var link = runs[index].getLinkUrl();
    if (link) {
      return link;
    }
  }

  return '';
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

function getHeaderIndex_(headers, aliases) {
  for (var index = 0; index < headers.length; index += 1) {
    if (aliases.indexOf(headers[index]) >= 0) {
      return index;
    }
  }

  return -1;
}

function parsePositiveNumber_(value) {
  var text = String(value || '').trim();
  if (!text) {
    return '';
  }

  var parsed = Number(text.replace(/%$/, ''));
  return isFinite(parsed) && parsed > 0 ? parsed : '';
}

function compactStyle_(style) {
  if (!style) {
    return null;
  }

  var compacted = {};
  if (style.backgroundColor) compacted.backgroundColor = style.backgroundColor;
  if (style.textColor) compacted.textColor = style.textColor;
  if (style.fontFamily) compacted.fontFamily = style.fontFamily;
  if (style.fontSize) compacted.fontSize = style.fontSize;
  if (style.bold === true) compacted.bold = true;
  if (style.italic === true) compacted.italic = true;
  if (style.underline === true) compacted.underline = true;
  if (style.strikethrough === true) compacted.strikethrough = true;

  return Object.keys(compacted).length > 0 ? compacted : null;
}

function getSideDefinitions_(spreadsheet, sideSheetName) {
  var sheet = spreadsheet.getSheetByName(sideSheetName);
  if (!sheet) {
    return {};
  }

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) {
    return {};
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
  var headers = values[0].map(normalizeHeader_);

  var idIndex = getHeaderIndex_(headers, ['id', 'key', 'side_id', 'number', 'no']);
  var textIndex = getHeaderIndex_(headers, ['text', 'label', 'content', 'title']);
  var heightIndex = getHeaderIndex_(headers, ['height', 'height_percent', 'percent', 'pct']);
  var backgroundColorIndex = getHeaderIndex_(headers, [
    'background_color',
    'bg_color',
    'background',
    'bg',
  ]);
  var textColorIndex = getHeaderIndex_(headers, [
    'text_color',
    'font_color',
    'foreground_color',
    'foreground',
  ]);
  var fontFamilyIndex = getHeaderIndex_(headers, ['font_family', 'font']);
  var fontSizeIndex = getHeaderIndex_(headers, ['font_size', 'text_size']);

  if (idIndex < 0 || textIndex < 0 || heightIndex < 0) {
    return {};
  }

  var definitions = {};

  for (var rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    var id = String(values[rowIndex][idIndex] || '').trim();
    if (!id) {
      continue;
    }

    var text = String(values[rowIndex][textIndex] || '').trim();
    var height = parsePositiveNumber_(values[rowIndex][heightIndex]);
    var textCellStyle = extractCellStyle_({
      backgroundColor: backgrounds[rowIndex][textIndex],
      textColor: fontColors[rowIndex][textIndex],
      fontFamily: fontFamilies[rowIndex][textIndex],
      fontSize: fontSizes[rowIndex][textIndex],
      fontWeight: fontWeights[rowIndex][textIndex],
      fontStyle: fontStyles[rowIndex][textIndex],
      fontLine: fontLines[rowIndex][textIndex],
    });
    var textRuns = extractRuns(richTextValues[rowIndex][textIndex], textCellStyle);
    var baseStyle = compactStyle_(textCellStyle) || {};

    if (backgroundColorIndex >= 0) {
      var backgroundColor = String(values[rowIndex][backgroundColorIndex] || '').trim();
      if (backgroundColor) {
        baseStyle.backgroundColor = backgroundColor;
      }
    }

    if (textColorIndex >= 0) {
      var textColor = String(values[rowIndex][textColorIndex] || '').trim();
      if (textColor) {
        baseStyle.textColor = textColor;
      }
    }

    if (fontFamilyIndex >= 0) {
      var fontFamily = String(values[rowIndex][fontFamilyIndex] || '').trim();
      if (fontFamily) {
        baseStyle.fontFamily = fontFamily;
      }
    }

    if (fontSizeIndex >= 0) {
      var fontSize = parsePositiveNumber_(values[rowIndex][fontSizeIndex]);
      if (fontSize) {
        baseStyle.fontSize = fontSize;
      }
    }

    definitions[id] = {
      id: id,
      text: text,
      textRuns: textRuns.length > 0 ? textRuns : undefined,
      height: height || undefined,
      backgroundColor: baseStyle.backgroundColor || undefined,
      textColor: baseStyle.textColor || undefined,
      fontFamily: baseStyle.fontFamily || undefined,
      fontSize: baseStyle.fontSize || undefined,
      bold: baseStyle.bold === true ? true : undefined,
      italic: baseStyle.italic === true ? true : undefined,
      underline: baseStyle.underline === true ? true : undefined,
      strikethrough: baseStyle.strikethrough === true ? true : undefined,
    };
  }

  return definitions;
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
