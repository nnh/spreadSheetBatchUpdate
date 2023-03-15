/**
 * @param {Object} batchUpdateRequest Request body.
 * @param {string} spreadsheetId Spreadsheet ID.
 * @return {Object} Sheet object.
 * @requires Sheets Api.
 */
function execBatchUpdate(batchUpdateRequest, spreadsheetId){
  return Sheets.Spreadsheets.batchUpdate(batchUpdateRequest, spreadsheetId).updatedSpreadsheet.sheets;
}
/**
 * Add "includeSpreadsheetInResponse: true" to the request body and return.
 * @param {Object} requests Request body.
 * @return {Object}
 */
function editBatchUpdateRequest(requests){
  return {
    requests: requests,
    includeSpreadsheetInResponse: true,
  }
}
/**
 * Obtain a sheet object from a sheet name.
 * @param {Object} sheets Sheet object.
 * @param {string} sheetName Name of the sheet to be extracted.
 * @return {Object} Sheet object.
 */
function getSheetBySheetName(sheets, sheetName){
  return sheets.filter(x => x.properties.title === sheetName)[0];
}
/**
 * Obtain a sheet object from a sheet id.
 * @param {Object} sheets Sheet object.
 * @param {string} sheetId Id of the sheet to be extracted.
 * @return {Object} Sheet object.
 */
function getSheetBySheetId(sheets, sheetId){
  return sheets.filter(x => x.properties.sheetId === sheetId)[0];
}
/**
 * Obtain a sheet id from a sheet name.
 * @param {Object} sheets Sheet object.
 * @param {string} sheetName Name of the sheet to be extracted.
 * @return {string} sheet id.
 */
function getSheetIdFromSheetName(sheets, sheetName){
  return sheets.map(x => x.properties.title === sheetName ? x.properties.sheetId : null).filter(x => x)[0];
}
/**
 * Set up a wrap around for all cells.
 * @param {string} sheetId sheet id.
 * @return {Object} Request body.
 */
function getAllCellWrapRequest(sheetId){
  return {
    'repeatCell': {
      'range': {'sheetId': sheetId},
      'cell': {
        'userEnteredFormat': {
          'wrapStrategy': 'WRAP',
          'verticalAlignment': 'TOP',
        },
      },
      'fields': 'userEnteredFormat.wrapStrategy,userEnteredFormat.verticalAlignment',
    }
  }
}
class BorderStyle{
  constructor(colorStyle=null){
    this.colorStyle = !colorStyle ? {
                                      'rgbColor': {
                                        'red': 0,
                                        'green': 0,
                                        'blue': 0,
                                        'alpha': 0,
                                      }, 
                                    }
                                  : colorStyle;
  }
  setBorderSolid(){
    return {
      'style': 'SOLID',
      'colorStyle': this.colorStyle,
    }
  }
  setBorderNone(){
    return {
      'style': 'NONE',
      'colorStyle': this.colorStyle,
    }
  }
}
function createBorderStyle(){
  return new BorderStyle();
}
function setBorderSolid(){
  throw new Error('Call this method after calling createBorderStyle.');
}
(function(global){
  function BorderStyle(){
    this.name = 'borderStyle';
  }
  BorderStyle.prototype.setBorderSolid = function(){
    return;
  }
  global.BorderStyle = BorderStyle;
})(this);
function setBorderNone(){
  throw new Error('Call this method after calling createBorderStyle.');
}
(function(global){
  function BorderStyle(){
    this.name = 'borderStyle';
  }
  BorderStyle.prototype.setBorderNone = function(){
    return;
  }
  global.BorderStyle = BorderStyle;
})(this);
/**
 * @param {string} sheetId sheet id.
 * @param {Object} rowCol {startRowIndex, startColumnIndex, endRowIndex, endColumnIndex}
 * @param {Object} top {style, width, color {}}
 */
function getUpdateBordersRequest(sheetId, rowCol, borders){
  const request = {
    'updateBorders': {
      'range': {
        'sheetId': sheetId,
        'startRowIndex': rowCol.startRowIndex,
        'endRowIndex': rowCol.endRowIndex,
        'startColumnIndex': rowCol.startColumnIndex,
        'endColumnIndex': rowCol.endColumnIndex,
      },
    }
  }
  if (borders.top){
    request.updateBorders.top = borders.top;
  }
  if (borders.bottom){
    request.updateBorders.bottom = borders.bottom;
  }
  if (borders.left){
    request.updateBorders.left = borders.left;
  }
  if (borders.right){
    request.updateBorders.right = borders.right;
  }
  if (borders.innerHorizontal){
    request.updateBorders.innerHorizontal = borders.innerHorizontal;
  }
  if (borders.innerVertical){
    request.updateBorders.innerVertical = borders.innerVertical;
  }
  return request;
}
/**
 * Set automatic row height settings.
 * @param {string} sheetId sheet id.
 * @param {number} startIndex start row index, ex.) A4 => 3.
 * @param {number} endIndex end row index, ex.) A7 => 6.
 * @return {Object} Request body.
 */
function getAutoResizeRowRequest(sheetId, startIndex, endIndex){
  return {
    'autoResizeDimensions': {
      'dimensions': {
        'sheetId': sheetId,
        'dimension': 'ROWS',
        'startIndex': startIndex,
        'endIndex' : endIndex,
      },
    }
  }
}
/**
 * Set the row height.
 * @param {string} sheetId sheet id.
 * @param {number} height the row height.
 * @param {number} startIndex start row index, ex.) A4 => 3.
 * @param {number} endIndex end row index, ex.) A7 => 6.
 * @return {Object} Request body.
 */
function getSetRowHeightRequest(sheetId, height=21, startIndex, endIndex){
  return {
    'updateDimensionProperties': {
      'range': {
        'sheetId': sheetId,
        'dimension': 'ROWS',
        'startIndex': startIndex,
        'endIndex' : endIndex,
      },
      'properties': {
        'pixelSize' : height,
      },
      'fields': 'pixelSize',
    }
  }
}
/**
 * Set the column width.
 * @param {string} sheetId sheet id.
 * @param {number} width the column width.
 * @param {number} startIndex start column index, ex.) B4 => 1.
 * @param {number} endIndex end column index, ex.) E19 => 4.
 * @return {Object} Request body.
 */
function getSetColWidthRequest(sheetId, width=120, startIndex, endIndex){
  return {
    'updateDimensionProperties': {
      'range': {
        'sheetId': sheetId,
        'dimension': 'COLUMNS',
        'startIndex': startIndex,
        'endIndex' : endIndex,
      },
      'properties': {
        'pixelSize' : width,
      },
      'fields': 'pixelSize',
    }
  }
}
function editNumberFormat(type='TEXT', pattern=null){
  const numberFormat = {};
  numberFormat.type = type;
  if (pattern){
    numberFormat.pattern = pattern;
  }
  const res = {
    'userEnteredFormat': {
      'numberFormat': numberFormat,
    }
  }
  return res;
}
/**
 * Sets the horizontal alignment of text in a cell.
 * @param {string} horizontalAlign 'LEFT' or 'CENTER' or 'RIGHT'.
 * @return {Object} Request body.
 */
function getHorizontalAlignmentRequest(horizontalAlign){
  const userEnteredFormat = {
    'userEnteredFormat': {
      'horizontalAlignment': horizontalAlign,
    }
  }
  return userEnteredFormat;
}
/**
 * Set the font to bold.
 * @param none.
 * @return {Object} Request body.
 */
function getFontBoldRequest(){
  const userEnteredFormat = {
    'userEnteredFormat': {
      'textFormat': {
        'bold': true,
      },
    }
  }
  return userEnteredFormat;
}
/**
 * Create and return a request body.
 * @param {string} sheetId sheet id.
 * @param {number} startRowIndex start row index, ex.) B4 => 3.
 * @param {number} startColumnIndex start column index, ex.) B4 => 1.
 * @param {string[][]} Values to be set in the cells.
 * @return {Object} Request body.
 */
function getRangeSetValueRequest(sheetId, startRowIndex, startColumnIndex, values){
  return { 
    'updateCells': {
      'range': getRangeGrid(sheetId, startRowIndex, startColumnIndex, values),
      'rows': editSetValues(values),
      'fields': 'userEnteredValue',
    }
  };
}
/**
 * Create and return a request body.
 * @param {string} sheetId sheet id.
 * @param {number} startRowIndex start row index, ex.) B4 => 3.
 * @param {number} startColumnIndex start column index, ex.) B4 => 1.
 * @param {number} endRowIndex end row index, ex.) B4 => 3.
 * @param {number} endColumnIndex end column index, ex.) B4 => 1.
 * @param {Object} userEnteredFormat
 * @param {string} fields ex.) 'userEnteredFormat.textFormat.bold'.
 * @return {Object} Request body.
 */
function getRangeSetFormatRequest(sheetId, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, userEnteredFormat, fields){
  const range = {
    'sheetId': sheetId,
    'startRowIndex': startRowIndex,
    'endRowIndex': endRowIndex + 1,
    'startColumnIndex': startColumnIndex,
    'endColumnIndex': endColumnIndex + 1,
  }
  let colArray = [];
  for (let col = startColumnIndex; col <= endColumnIndex; col++){
    colArray.push(userEnteredFormat);
  }
  let values = [];
  for (let row = startRowIndex; row <= endRowIndex; row++){
    values = [...values, {'values': colArray}];
  }
  return { 
    'updateCells': {
      'range': range,
      'rows': values,
      'fields': fields,
    }
  };
}
/**
 * Create and return a request body.
 * @param {string} sheetId sheet id.
 * @param {string} title Sheet name to be set.
 * @return {Object} Request body.
 */
function editRenameSheetRequest(sheetId, title){
  return {
    'updateSheetProperties': {
      'properties': {
        'sheetId': sheetId,
        'title': title,
      },
      'fields': 'title',
    },
  }
}
/**
 * Returns the type of value to be set.
 * @param {string[][]} Values to be set in the cells.
 * @return {string[][]} Request body.
 */
function editSetValues(testValues){
  const arr = testValues.map(row => {
    const cols = row.map(col => {
      const obj = {};
      obj.userEnteredValue = {};
      col = col === null ? '' : col;
      const type = col === true || col === false ? 'boolValue' 
                   : Number.isFinite(col) ? 'numberValue'
                   : toString.call(col) === '[object Date]' ? 'numberValue'
                   : col.substring(0, 1) === '=' ? 'formulaValue'
                   : 'stringValue';
      obj.userEnteredValue[type] = col;  
      return obj;
    });
    const values = {};
    values.values = cols;
    return values;
  });
  return arr;
}
/**
 * Set a range.
 * @param {string} sheetId sheet id.
 * @param {number} width the column width.
 * @param {number} startIndex start column index, ex.) B4 => 1.
 * @param {number} endIndex end column index, ex.) E19 => 4.
 * @return {Object} Request body.
 */
function getRangeGrid(sheetId, startRowIndex, startColumnIndex, values){
  const endRowIndex = startRowIndex + values.length;
  const endColumnIndex = startColumnIndex + values[0].length;
  return {
    'sheetId': sheetId,
    'startRowIndex': startRowIndex,
    'endRowIndex': endRowIndex,
    'startColumnIndex': startColumnIndex,
    'endColumnIndex': endColumnIndex
  }
}
/**
 * @param {number} spreadsheetId
 * @param {number} sheetId
 * @param {string} title Sheet name to be set.
 * @return {Object} Sheet object.
 * @requires Sheets Api.
 */
function renameSheet(spreadsheetId, sheetId, title){
  const batchUpdateRequest = 
  {
    'requests': [
      {
        'updateSheetProperties': {
          'properties': {
            'sheetId': sheetId,
            'title': title,
          },
          'fields': 'title',
        },
      },
    ],
    'includeSpreadsheetInResponse': true,
  };
  const sheets = Sheets.Spreadsheets.batchUpdate(batchUpdateRequest, spreadsheetId).updatedSpreadsheet.sheets;
  return sheets.filter(x => x.properties.title === title)[0];
}
/**
 * @param {number} spreadsheetId
 * @param {string} range ex.) 'sheet1!A2:B3'
 * @param {string} valueRenderOption 'FORMATTED_VALUE' or 'UNFORMATTED_VALUE' or 'FORMULA'.
 * @return {Object} range values.
 * @requires Sheets Api.
 */
function rangeGetValue(spreadsheetId, range, valueRenderOption='FORMATTED_VALUE'){
  const param = 
  {
    ranges: range,
    valueRenderOption: valueRenderOption
  };
  const values = Sheets.Spreadsheets.Values.batchGet(spreadsheetId, param);
  return values.valueRanges;
}
/**
 * @param {number} spreadsheetId
 * @param {Object} range
 * @param {string} valueRenderOption 'FORMATTED_VALUE' or 'UNFORMATTED_VALUE' or 'FORMULA'.
 * @return {string} range value.
 */
function getValueRangesValue(spreadsheetId, range, valueRenderOption='FORMATTED_VALUE'){
  const valueRanges = rangeGetValue(spreadsheetId, range, valueRenderOption);
  return valueRanges[0].values[0][0];
}
