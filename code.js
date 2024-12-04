function createTemplate() {
  var arr = [["TEMPLATE TITLE", "Probability", "Impact"], ["Item 1", 2, 2], ["Item 2", 3, 2], ["Item 4", 1, 2], ["Item 5", 1, 2]]

 var sheet = createRiskMappingSheet_();

  sheet.getRange(1, 1, arr.length, arr[0].length).setValues(arr);
};


function calculateCoordinates(range) {
  var data = sourceSheet.getRange(range).getValues();

  // Step 1: Read data, skipping the header row
  var dataValues = data.slice(1).map((row, index) => ({
    id: row[0],  // ID (from column 1)
    xAxis: row[1],  
    yAxis: row[2],
  }));

  var maxRange = getMaxGraphSize_(dataValues); // Gets the largest number in either column. This will define the size of the grid in the graph
  var quadrantSize = 1; // This defines the size of each quadrant (by default it should be 1)

  // Step 2: Group projects into quadrants
  const quadrantMap = {};

  // Loop through each project in the list and assign it to a quadrant 
  dataValues.forEach(project => {
    if (project.id === "" || project.xAxis === null) {
      return; // Skip this iteration
    };

    var key = `Quadrant [${project.xAxis},${project.yAxis}]`;
    if (!quadrantMap[key]) {
      quadrantMap[key] = {
        'xAxis': project.xAxis,
        'xStarting': project.xAxis - quadrantSize / 2,
        'yAxis': project.yAxis,
        'yStarting': project.yAxis - quadrantSize / 2,
        'points': []
      };
    };
    quadrantMap[key].points.push(project.id);
  });


  // Step 3: Adjust coordinates for points in each quadrant 
  const adjustedData = [data[0]]; // Header row
  adjustedData[0].push("Colour");

  let minValue = Infinity;
  let maxValue = -Infinity;

  for (const key in quadrantMap) {
    var points = quadrantMap[key]["points"];
    var pointsInQuadrant = points.length;

    var offsetFactors = findNearPerfectSquare_(pointsInQuadrant);
    var offsetY = quadrantSize / ((2 * offsetFactors[1]));
    var offsetX = quadrantSize / ((2 * offsetFactors[0]));

    for (var i = 0; i < points.length; i++) {
      var floorX = (Math.floor((i / offsetFactors[1])) + 0.5) * 2;
      var floorY = ((i % offsetFactors[1]) * 2) + 1;

      var positionX = quadrantMap[key]["xStarting"] + (floorX) * offsetX;
      var positionY = quadrantMap[key]["yStarting"] + (floorY) * offsetY;

      var colourValue = quadrantMap[key]["xAxis"] * quadrantMap[key]["yAxis"];

      adjustedData.push([points[i], positionX, positionY, colourValue]);
    }
  };

  // Step 4: Check if "Adjusted Data" sheet exists; if not, create it
  var name = sourceSheet.getSheetName();
  var adjustedSheet = ss.getSheetByName(name + " (graph data)");
  if (!adjustedSheet) {
    adjustedSheet = ss.insertSheet(name + " (graph data)");
  };

  adjustedSheet.clear();

  // Step 5: Write adjusted data to the sheet
  var outputRange = adjustedSheet.getRange(1, 1, adjustedData.length, adjustedData[0].length);
  outputRange.setValues(adjustedData);

  var chartOptions = {
    "title": adjustedData[0][0],
    "hAxis": {
      "title": adjustedData[0][1],
      "minValue": 0.5,
      "maxValue": maxRange + 0.5,
      "step": 1
    },
    "vAxis": {
      "title": adjustedData[0][2],
      "minValue": 0.5,
      "maxValue": maxRange + 0.5,
      "step": 1
    }
  };

  var chartRange = adjustedSheet.getRange(2, 1, adjustedData.length - 1, 3); // Exclude header
  return {
    "chartRange": chartRange,
    "chartOptions": chartOptions,
    "adjustedData": adjustedData
  };
};

function createChart(chartRange, chartOptions) {

  // Create the chart
  const chartBuilder = sourceSheet.newChart()
    .setChartType(Charts.ChartType.BUBBLE)
    .addRange(chartRange) // Include the "Colour" column in this range
    .setPosition(5, 5, 0, 0)
    .setOption('hAxis', { 
      title: chartOptions.hAxis.title, 
      minValue: chartOptions.hAxis.minValue, 
      maxValue: chartOptions.hAxis.maxValue, 
      gridlines: { step: chartOptions.hAxis.step }
    })
    .setOption('vAxis', { 
      title: chartOptions.vAxis.title, 
      minValue: chartOptions.vAxis.minValue, 
      maxValue: chartOptions.vAxis.maxValue, 
      gridlines: { step: chartOptions.vAxis.step }
    })
    .setOption('bubble', {
      opacity: 0.7,
      sizeAxis: { 
        minSize: 5, 
        maxSize: 30 
      }
    })
    .setOption('legend', { position: 'none' })
    .setOption('width', 600)
    .setOption('height', 600);

    /* Set colours. I never got this to work.  
    .setOption('colorAxis', {
      // Ensure this is a numerical range that corresponds to the "Colour" column
      minValue: chartOptions.hAxis.minValue || 0, // Default min value
      maxValue: chartOptions.hAxis.maxValue || 16, // Default max value
      colors: ['#FFA500', '#008000', '#0000FF'] // Your specified color gradient (Orange -> Green -> Blue)
    }); */
  var newChart = chartBuilder.build();
  sourceSheet.insertChart(newChart);
};






