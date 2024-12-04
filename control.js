function runAll() {
  var dataRange = showRangePicker_();

  var coordinates = calculateCoordinates(dataRange);

  createChart(coordinates.chartRange, coordinates.chartOptions, coordinates.adjustedData); 
};

function onOpen(e) {
  var ui = SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Set template', 'createTemplate')
    .addItem('Create graph', 'runAll')
    .addToUi(); 
};

/* ---------------------------------- */
/* --------- Global Variables -------- */
/* ---------------------------------- */

const ss = SpreadsheetApp.getActiveSpreadsheet(); 
const sourceSheet = ss.getActiveSheet();
var docProps = PropertiesService.getDocumentProperties();

/* --------------------------------------- */
/* --------- Supporting Functions -------- */
/* --------------------------------------- */

function findNearPerfectSquare_(n) {
    // Start with the integer square root of n
    var output = []
    let a = b = Math.floor(Math.sqrt(n));

    while (a * b < n) {
      if (a === b) {
        a++
      } else {
        b++
      };
    };
    return [a,b];
};

function showRangePicker_() {
  const ui = SpreadsheetApp.getUi();

  var lastRange = docProps.getProperty('lastRange');

  var msg;
  if (lastRange) {
    msg = `\n\nLast used range was: "${lastRange}". Press ok to reuse.`
  } else {
    msg = "" };

  // Prompt the user to input a range
  const result = ui.prompt(
    'Range Picker',
    `Please enter the range (e.g., A1:D10): ${msg}`,
    ui.ButtonSet.OK_CANCEL
  );

  // Process the input
  if (result.getSelectedButton() === ui.Button.OK) {
    var rangeString = result.getResponseText();
    if (rangeString === null || rangeString === "") {
      rangeString = lastRange
    };

    let range;
    try {
      range = ss.getRange(rangeString); // Validate the range
      docProps.setProperty('lastRange', rangeString);

      return rangeString;
    } catch (e) {
      ui.alert('Invalid range. Please try again.');
    }
  } else {
    ui.alert('Range selection canceled.');
    return;
  };
};

function createRiskMappingSheet_() {
  let sourceSheet = ss.getSheetByName("Risk Template");
  if (!sourceSheet) {
    sourceSheet = ss.insertSheet("Risk Template");
  };

  return sourceSheet;
};

function getMaxGraphSize_(data) {
  const maxValue = data.reduce((max, project) => {
    return Math.max(max, project.xAxis, project.yAxis);
  }, -Infinity);

  // Check if maxValue is valid
  if (maxValue === -Infinity) {
    Logger.log("Error: maxValue calculation failed");
    return 0; // Default value for maxRange
  }

  return maxValue;
};





