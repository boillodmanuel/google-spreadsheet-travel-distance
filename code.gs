/** DOC
 *  
 * SETUP:
 *  1. set the current period
 *  2. run setupDataValidation
 *  3. when doc is edited, routes and kms are updated
 *  
 *  To start a new year, duplicates google sheets, archive and protect it and restart SETUP above
 */

// Current period
var CURRENT_YEAR = 2019
var FIRST_DATE = new Date(CURRENT_YEAR, 6, 9)
var LAST_DATE = new Date(CURRENT_YEAR, 11, 31)

// Constants
var FIRST_ROW = 3
var DATE_ADDRESS_COL = 1
var ORIGIN_NAME_COL = 2
var ORIGIN_ADDRESS_COL = 3
var DESTINATION_NAME_COL = 4
var DESTINATION_ADDRESS_COL = 5
var ROUND_TRIP_COL = 6
var GOOGLE_ROUTE_KM_COL = 8
var GOOGLE_ROUTE_ORIGIN_COL = 9
var GOOGLE_ROUTE_DEST_COL = 10
var TOTAL_KM_COL = 11
var MESSAGE_COL = 12
var BIG_TOTAL_KM_COL = 12
var TOTAL_KM_COL_LETTER = 'K'

var COLOR_INFO = null
var COLOR_WARN = '#ff9900'
var COLOR_ERROR = '#ff0000'
    
var spreadsheet = SpreadsheetApp.getActive();

var addresses = spreadsheet.getSheetByName("Adresses");
var addressesLastRow = addresses.getLastRow();

var km = spreadsheet.getSheetByName("KM " + CURRENT_YEAR);
var kmLastRow = km.getLastRow();
var debugMessage = km.getRange('K2');

// DEBUG
var DEBUG_ENABLED = false

// Sheet Setup (to execute one)
function setupDataValidation() {
  km.getRange(1, 1, km.getLastRow(), km.getLastColumn()).clearDataValidations();
  
  // Date validation
  km.getRange(FIRST_ROW, DATE_ADDRESS_COL, kmLastRow - FIRST_ROW).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .setAllowInvalid(false)
      .setHelpText('Saisissez une date entre ' + getFormattedDate(FIRST_DATE) + ' et ' + getFormattedDate(LAST_DATE))
      .requireDateBetween(FIRST_DATE, LAST_DATE)
      .build());
  
  // Addresses name
  km.getRange(FIRST_ROW, ORIGIN_NAME_COL, kmLastRow - FIRST_ROW + 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .setAllowInvalid(true)
      .requireValueInRange(spreadsheet.getRange('Adresses!$A$2:$A$' + addressesLastRow), true)
      .build());
  km.getRange(FIRST_ROW, DESTINATION_NAME_COL, kmLastRow - FIRST_ROW + 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .setAllowInvalid(true)
      .requireValueInRange(spreadsheet.getRange('Adresses!$A$2:$A$' + addressesLastRow), true)
      .build());
  
  // Cells Format
  km.getRange(FIRST_ROW, 1, kmLastRow - FIRST_ROW + 1, km.getLastColumn())
    .setFontFamily('Arial')
    .setFontSize(10)
    .setFontColor(null)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top');
  km.getRange(FIRST_ROW, ORIGIN_ADDRESS_COL, kmLastRow - FIRST_ROW + 1)
    .setFontFamily('Arial')
    .setFontSize(8)
    .setFontColor(null);
  km.getRange(FIRST_ROW, DESTINATION_ADDRESS_COL, kmLastRow - FIRST_ROW + 1)
    .setFontFamily('Arial')
    .setFontSize(8)
    .setFontColor(null);
  km.getRange(FIRST_ROW, ROUND_TRIP_COL, kmLastRow - FIRST_ROW + 1)
    .setHorizontalAlignment('center');
  km.getRange(FIRST_ROW, GOOGLE_ROUTE_ORIGIN_COL, kmLastRow - FIRST_ROW + 1)
    .setFontFamily('Arial')
    .setFontSize(8)
    .setFontColor(null);
  km.getRange(FIRST_ROW, GOOGLE_ROUTE_DEST_COL, kmLastRow - FIRST_ROW + 1)
    .setFontFamily('Arial')
    .setFontSize(8)
    .setFontColor(null);
  
  // Round trip checkbox
  km.getRange(FIRST_ROW, ROUND_TRIP_COL, kmLastRow - FIRST_ROW + 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .setAllowInvalid(false)
      .requireCheckbox()
      .build());
  
  // YEAR TOTAL
  km.getRange(FIRST_ROW - 2, BIG_TOTAL_KM_COL).setValue('TOTAL KMS ' + CURRENT_YEAR);
  km.getRange(FIRST_ROW - 1, BIG_TOTAL_KM_COL).setFormula('=SUM(' + TOTAL_KM_COL_LETTER + FIRST_ROW + ':' + TOTAL_KM_COL_LETTER + kmLastRow + ')');
};

/** Triggered automatically */
function onEdit(e) {
  var range = e.range;
  var col = range.getColumn()
  var row = range.getRow()
  
  if (DEBUG_ENABLED) debug('on edit ' + range.toString() + col + row)
  
  if (row >= FIRST_ROW) {
    if (col === ORIGIN_NAME_COL) {
      updateOriginAddress(row);
    } else if (col === DESTINATION_NAME_COL) {
      updateDestinationAddress(row);
    } else if (col === ORIGIN_ADDRESS_COL) {
      km.getRange(row, ORIGIN_NAME_COL).clearContent()
      updateRoute(row);
    } else if (col === DESTINATION_ADDRESS_COL) {
      km.getRange(row, DESTINATION_NAME_COL).clearContent()
      updateRoute(row);
    } else if (col === ROUND_TRIP_COL) {
      updateRoundTrip(row);
    }
  }
}

function updateOriginAddress(row) {
  updateAddress(row, ORIGIN_NAME_COL, ORIGIN_ADDRESS_COL)
}
function updateDestinationAddress(row) {
  updateAddress(row, DESTINATION_NAME_COL, DESTINATION_ADDRESS_COL)
}
function updateAddress(row, nameCol, addressCol) {
  var nameRange = km.getRange(row, nameCol);
  if (nameRange.isBlank()) {
    km.getRange(row, addressCol).clearContent()
  } else {
    var address = addressLookup(nameRange.getValue());
    if (address == null) {
      error(row, 'Erreur : addresse non trouvée dans carnet d\'addresse');
      return
    } else {
      km.getRange(row, addressCol).setValue(address)
    }
  }
  updateRoute(row);
}

function validateData(row) {
  // Data
  var originRange = km.getRange(row, ORIGIN_ADDRESS_COL);
  var destinationRange = km.getRange(row,DESTINATION_ADDRESS_COL);
  
  // Validate data  
  if (!originRange.isBlank() && destinationRange.isBlank()) {
    warn(row, "Adresse destination manquante")
    return false;
  }
  if (originRange.isBlank() && !destinationRange.isBlank()) {
    warn(row, "Adresse d'origine manquante")
    return false;
  }
  if (originRange.isBlank() && destinationRange.isBlank()) {
    clearMessage(row)
    return false;
  }
  return true;
}

function updateRoute(row) {
  info(row, 'updating...')
  if (DEBUG_ENABLED) debug('updateRoute ' + row)
  
  // Clear output
  km.getRange(row, GOOGLE_ROUTE_KM_COL, 1, 4).clearContent();
  
  if (!validateData(row)) return;

  // Get route
  var route = getRoute(km.getRange(row, ORIGIN_ADDRESS_COL).getValue(), km.getRange(row,DESTINATION_ADDRESS_COL).getValue());
  if (route === null) {
    error(row, 'Erreur : addresse inconnue');
    return;
  }
  
  // Display route
  var distance = Math.round(route.distance.value / 1000);
  
  km.getRange(row, GOOGLE_ROUTE_KM_COL).setValue(distance);
  km.getRange(row, GOOGLE_ROUTE_ORIGIN_COL).setValue(route.start_address);
  km.getRange(row, GOOGLE_ROUTE_DEST_COL).setValue(route.end_address);
  
  var distanceText = route.distance.text.toString();
  if (!distanceText.endsWith(' km') && !distanceText.endsWith(' m')) {
    error(row, "Erreur: Unité de distance à vérifier - valeur google = '" + distanceText + "'");
    return;
  }
  
  doUpdateRoundTrip(row);
  clearMessage(row)
}

function updateRoundTrip(row) {
  if (km.getRange(row, GOOGLE_ROUTE_KM_COL).isBlank()) return;
  
  info(row, 'updating...')
  if (DEBUG_ENABLED) debug('updateRoundTrip ' + row)
  
  // Clear output
  km.getRange(row, TOTAL_KM_COL).clearContent();
  
  if (!validateData(row)) return;
  
  doUpdateRoundTrip(row);
  clearMessage(row)
}

function doUpdateRoundTrip(row) {
  var roundTrip = km.getRange(row,ROUND_TRIP_COL).getValue()
  var distance = km.getRange(row, GOOGLE_ROUTE_KM_COL).getValue()
  
  totalDistance = roundTrip ? distance * 2 : distance;
  km.getRange(row, TOTAL_KM_COL).setValue(totalDistance);
}

function rowMessage(row, message, color) {
  km.getRange(row, MESSAGE_COL)
       .setValue(message)
       .setFontColor(color);
}
function info(row, message) {
    rowMessage(row, message, COLOR_INFO);
}
function warn(row, message) {
    rowMessage(row, message, COLOR_WARN);
}
function error(row, message) {
    rowMessage(row, message, COLOR_ERROR);
}
function clearMessage(row) {
    rowMessage(row, "", COLOR_INFO);
}

function debug(message) {
  debugMessage.setValue('debug: ' + message + '\n' + debugMessage.getValue());
}

function logRoute(route) {
  Logger.log("route.start_address:  " + route.start_address);
  Logger.log("route.end_address:    " + route.end_address);
  Logger.log("route.distance.text:  " + route.distance.text);
  Logger.log("route.distance.value: " + route.distance.value);
}

function getRoute(source, destination) {
  var directions = Maps.newDirectionFinder()
    .setOrigin(source)
    .setDestination(destination)
    .setMode(Maps.DirectionFinder.Mode.DRIVING)
    .getDirections();
  
  //Logger.log(directions);
  
  if (directions.routes.length) {
    return directions.routes[0].legs[0];
  } else {
    return null;
  }
}

// Polyfill
if (!String.prototype.endsWith) {
	String.prototype.endsWith = function(search, this_len) {
		if (this_len === undefined || this_len > this.length) {
			this_len = this.length;
		}
		return this.substring(this_len - search.length, this_len) === search;
	};
}

function getFormattedDate(date) {
  var year = date.getFullYear();
  var month = (1 + date.getMonth()).toString();
  month = month.length > 1 ? month : '0' + month;
  var day = date.getDate().toString();
  day = day.length > 1 ? day : '0' + day;
  return day + '/' + month + '/' + year;
}


function addressLookup(name) {
  var data = addresses.getRange(1,1, addressesLastRow, 2).getValues();// create an array of data from columns A and B
  for(nn=0;nn<data.length;++nn){
    if (data[nn][0]=== name) {
      return data[nn][1];
    }
  }
  return null
}

// Manual testing
  
function testUpdateFirstRow() {
  updateRoute(FIRST_ROW)
}

function testAddressLookup() {
  address = addressLookup('Domicile')
  Logger.log(address)
}


function testGetRouteError() {
  getRoute("nantes", "toto")
}
function testGetRouteSmall() {
  getRoute("31 rue Michel Chauty 44800 St-Herblain", "35 rue Michel Chauty 44800 St-Herblain")
}
