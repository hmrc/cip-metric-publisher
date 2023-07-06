var SERIES_COLORS = {0: {color: "00008B"},
                     1: {color: "008d8d"},
                     2: {color: "fbbc04"}}
var DEFAULT_CHART_NUM_MONTHS = 3
var DEFAULT_GRAPH_TYPE = "Line"     

//
// Searches the current sheet for a graph of a particular name
//
function searchGraphByName (searchName) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    var charts = sheet.getCharts()
    var apiCharts = Sheets.Spreadsheets.get(ss.getId(), { ranges: [sheet.getSheetName()], fields: "sheets(charts)" }).sheets[0].charts;

    for (var i = 0; i < apiCharts.length; i++) {
        chart = apiCharts[i]
        var { altText } = chart.spec;
        if (altText) {
            // Check whether the alt text graph name matches the search name
            var regExp = new RegExp('.*(Graph): ([^\\s]+)', "i"); 
            var match = regExp.exec(altText);
            if (match) {
                graphAltName = match[2]
                if (graphAltName === searchName) {
                    return (charts[i])
                }
            }
        }
    }
    // Not found
    return null
}

// Generates a chart image Blob, taking:
//   configiration object, generated by deriveMetricsConfiguration function
//   sourceIds - an array of source IDs representing the graph series
//   graphConfig - a JSON formatted string containing configuration for the graph itself. Option names/values correspond to
//                 Google chart options, Example:
//                  {"numberOfMonths": <number of months history to chart>,
//                   "options": [{"option name": <option value>},{"option2 name": <option2 value>},...]}

function createGraphBlob (configuration, sourceIds, chartConfigString) {
  var chartConfig = JSON.parse(chartConfigString)

  var numMonths = DEFAULT_CHART_NUM_MONTHS
  if (chartConfig.numberOfMonths) {var numMonths = chartConfig.numberOfMonths}
  var graphType = DEFAULT_GRAPH_TYPE
  if (chartConfig.graphType) {graphType = chartConfig.graphType}
  var isStacked = false
  if (chartConfig.isStacked) {var isStacked = chartConfig.isStacked}

  // Construct the graph structure, consisting of month along the X axis, and a series per data source
  var data = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, 'Month')
  for (i = 0; i < sourceIds.length; i++) {
    data = data.addColumn(Charts.ColumnType.NUMBER, configuration["nameMap"][sourceIds[i]])
  }

  // Retrieve the title row from the datasheet tab
  titlesRow = configuration["titles"]

  // Loop over the months of data, constructing the data series
  for (mc = titlesRow.length-numMonths; mc < titlesRow.length; mc++) {
    monthTitle = titlesRow[mc]
    monthGraphData = [monthTitle]
//    monthGraphData.push(monthTitle)
    // Loop over the sourceIds and generate the datapoints for the month
    for (s = 0; s < sourceIds.length; s++) {
      sourceId = sourceIds[s]
      sourceIdData = configuration["historyMap"][sourceId][mc]//[titlesRow.length-numMonths,titlesRow.length]
      monthGraphData.push(Number(sourceIdData))
    }
    // Add the months datapoints to the graph
    data = data.addRow(monthGraphData)
  }
  data = data.build();

  // var view = Charts.newDataViewDefinition().setColumns([0,1,{ calc: "stringify",
  //                        sourceColumn: 1,
  //                        type: "string",
  //                        role: "annotation" }]);

  switch (graphType) {
    case "Line":
      var chart = Charts.newLineChart()
      break;
    case "Bar":
      var chart = Charts.newBarChart()
      if (isStacked) {chart = chart.setStacked()}
      break;
    case "Column":
      var chart = Charts.newColumnChart()
      if (isStacked) {chart = chart.setStacked()}
      break;
    default:
      var chart = Charts.newLineChart()
      break;
  }
  chart = chart.setDataTable(data)
      .setDimensions(1048, 768)
      .setOption('series', SERIES_COLORS)
//.setDataViewDefinition(view)

  if (chartConfig.options) {
    for(var optionKey in chartConfig.options){
      Logger.log("Option " + optionKey) 
      Logger.log("Value " + JSON.stringify(chartConfig.options[optionKey]))
      chart = chart.setOption(optionKey, chartConfig.options[optionKey])
    }
  }

  // Construct, render and return Blob
  var chart = chart.build();
  var imageBlob = chart.getAs('image/png');
  return imageBlob;
}
