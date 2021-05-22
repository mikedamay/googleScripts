/***
 * This script must be embedded in the MDAM-EXES... spreadsheet
 * TODO: make start and end rows dynamic
 * TODO: assert that monthly data range matches array
 *
 */
var COL_DATE = 1;
var COL_DESCRIPTION = 2;
var COL_PURCHASE_TYPE = 3;
var COL_RECEIPTS = 4;
var COL_PAYMENTS = 5;
var COL_CASH = 6;
var COL_GROCERIES = 7;
var COL_FLAT_N_PERSONAL = 8;
var COL_ENTERTAINMENT_1 = 9;
var COL_PHONES_N_COMPUTERS = 10;
var COL_RECREATION = 11;
var COL_SUNDRY = 12;
var COL_INVESTMENTS = 13;
var COL_ENTERTAINMENT_2 = 14;
var COL_BARCLAYCARD_RECEIPTS = 15; // not used
var COL_CREDIT_CARD_PAYMENTS = 16;

var COL_FIRST = COL_DATE;
var COL_LAST = COL_CREDIT_CARD_PAYMENTS;

// the row order of the monthly data mirrors the column order of the main tables
var ROFFSET_CASH = COL_CASH - COL_CASH
var ROFFSET_GROCERIES = COL_GROCERIES - COL_CASH
var ROFFSET_FLAT_N_PERSONAL = COL_FLAT_N_PERSONAL - COL_CASH
var ROFFSET_ENTERTAINMENT_1 = COL_ENTERTAINMENT_1 - COL_CASH
var ROFFSET_PHONES_N_COMPUTERS = COL_PHONES_N_COMPUTERS - COL_CASH
var ROFFSET_RECREATION = COL_RECREATION - COL_CASH
var ROFFSET_SUNRY = COL_SUNDRY - COL_CASH
var ROFFSET_INVESTMENTS = COL_INVESTMENTS - COL_CASH
var ROFFSET_ENTERTAINMENT_2 = COL_ENTERTAINMENT_2 - COL_CASH

var MONTHLY_DATA_NUM_ROWS = 9;
var MONTHLY_DATA_NUM_COLS = 12;

function go() {
    doMonthlyAccumulation(SpreadsheetApp.getActiveSheet(), findStartAndEndForPaymentMethods());
}

function doMonthlyAccumulation(sheet, paymentMethodRanges) {
    var barclayCardAccumulator = accumulate(sheet
        , paymentMethodRanges.barclayCardStartRow, paymentMethodRanges.barclayCardEndRow);
    var postOfficeAccumulator = accumulate(sheet
        , paymentMethodRanges.postOfficeStartRow, paymentMethodRanges.postOfficeEndRow);
    var bankAccumulator = accumulate(sheet
        , paymentMethodRanges.bankStartRow, paymentMethodRanges.bankEndRow);
    var accumulator = getEmptyAccumulator();
    addRangeToAccumulator(accumulator, barclayCardAccumulator);
    addRangeToAccumulator(accumulator, postOfficeAccumulator);
    addRangeToAccumulator(accumulator, bankAccumulator);
    createMonthlyDataTable(sheet, accumulator);
}

function createMonthlyDataTable(sheet, accumulator) {
    var topLeft = findMonthDataTopLeft(sheet);
    var rng = sheet.getRange(topLeft.top, topLeft.left, MONTHLY_DATA_NUM_ROWS, MONTHLY_DATA_NUM_COLS);
    rng.setValues(accumulator);
}

function accumulate(sheet, startRow, endRow) {
    var coffset_cash = COL_CASH - 1;  // sheet grid is numbered from 1, equivalent data values from 0
    var accumulator = getEmptyAccumulator();
    var rng = sheet.getRange(startRow, COL_FIRST, endRow, COL_LAST);
    var data = rng.getValues();
    for (ii = 0; ii < endRow - startRow; ii++) {
        var dt = data[ii][COL_DATE - 1];
        if (typeof dt !== "object") {
            continue;
        }
        var month = dt.getMonth();
        for (jj = coffset_cash; jj <= COL_ENTERTAINMENT_2 - 1; jj++) {
            if (!isNumeric(data[ii][jj])) {
                continue;
            }
            Logger.log(ii + ", " + jj + ": " + data[ii][jj]);
            accumulator[jj - coffset_cash][month] += parseFloat(data[ii][jj]);
        }
    }
    return accumulator;
}

function addRangeToAccumulator(aggregateAccumulator, accumulator)
{
    for (ii = 0; ii < aggregateAccumulator.length; ii++) {
        for (jj = 0; jj < aggregateAccumulator[ii].length; jj++) {
            aggregateAccumulator[ii][jj] += accumulator[ii][jj];
        }
    }
}

function convertDates() {
    var sheet = SpreadsheetApp.getActiveSheet();
    for (ii = 1; ii < 1100; ii++ ) {
        convertDate(sheet, ii);
    }
}


function convertDate(sheet, row) {
    var rng = sheet.getRange(row, COL_DATE);
    var dtNew = false;
    var dt = rng.getValue();
    if (typeof dt === "object")
    {
        Logger.log("object found " + dt.getFullYear() + " " + dt.getDate() + " " + dt.getMonth());
        dtNew = new Date(dt.getFullYear(), dt.getDate() - 1, dt.getMonth() + 1);
    }
    else if (typeof dt === "string")
    {
        Logger.log("string found");
        dtNew = dateFromString(dt)
    }
    if (dtNew)
    {
        rng.setValue(dtNew)
    }
}

function findStartAndEndForPaymentMethods(sheet) {
    //sheet.getParent().setSpreadsheetLocale("en_GB");
    //Logger.log("locale: " + sheet.getParent().getSpreadsheetLocale());
    //var data = sheet.getDataRange().getValues();
    /*
    for (var i = 0; i < data.length; i++) {
      Logger.log(i +": " + data[i][COL_DATE] + " " + typeof(data[i][COL_DATE]));
    }
    */
    //data[19][COL_DATE] = new Date("01/13/2021");

    var paymentMethodRanges = {
        barclayCardStartRow : 2 + 1, barclayCardEndRow : 862 - 1
        ,bankStartRow : 875 + 1, bankEndRow : 1064 - 1
        ,postOfficeStartRow : 1066 + 1, postOfficeEndRow : 1077 - 1};
    return paymentMethodRanges;
}

function findMonthDataTopLeft(sheet) {
    var monthTableAnchor = {top : 1214 + 1, left : 5 + 2};
    return monthTableAnchor;
}

function dateFromString(str) {
    var dayOfMonth = str.slice(0, 2);
    var month = str.slice(3, 5);
    var year = str.slice(6, 10);
    if (!isNumeric(dayOfMonth))
    {
        Logger.log("it's not a number " + year + " " + month + " " + dayOfMonth);
        return false;
    }
    var dt = new Date(year, month - 1, dayOfMonth);
    return dt;
}

// https://stackoverflow.com/a/1830844/96167
function isNumeric(n) {
    return !isNaN(parseFloat(n)) && isFinite(n);
}

function getEmptyAccumulator() {
    return [
        [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        ,[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        ,[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        ,[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        ,[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        ,[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        ,[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        ,[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        ,[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    ];
}

function displayAccumulator(accumulator) {
    for (ii = 0; ii < accumulator.length; ii++ ) {
        var line = "";
        for (jj = 0; jj < accumulator[ii].length; jj++) {
            line += accumulator[ii][jj] + " ";
        }
        Logger.log(line);
    }
}

