var METRIC_ID_CELL_TEXT = "Metric ID"
var ACCOUNTSSERVICES_CELL_TEXT = "Accounts/Services"
var METRIC_TYPE_CELL_TEXT = "Metric Type"
var QUERY_METRIC_CELL_TEXT = "Query Metric"
var FILTER_CELL_TEXT = "Filter"

function getColumnDataByName(sheet, columnName) {
  var location = sheet.createTextFinder(columnName).matchEntireCell(true).findNext();
  if (!location)  {
      throw "Couldn't find " + columnName + " marker, did you delete a cell called '" + columnName + "'?"
  }
  var columnData   = sheet.getRange(location.getRow()+1,location.getColumn(),sheet.getLastRow()-location.getRow()).getDisplayValues()
  return columnData
}

function cipMetrics_queryMetricsForColumn() {
  sheet = SpreadsheetApp.getActiveSheet()
  
  startDate = new Date(sheet.getRange(1,sheet.getActiveCell().getColumn(),1).getValues()[0][0]);
  Logger.log(startDate)
  if (isNaN(startDate.getFullYear())) {
      throw "Couldn't find a valid report publish date to use"
  }

  var metricTypeColumn  = getColumnDataByName(sheet, METRIC_TYPE_CELL_TEXT)
  var keysColumn        = getColumnDataByName(sheet, METRIC_ID_CELL_TEXT)
  var queryMetricColumn = getColumnDataByName(sheet, QUERY_METRIC_CELL_TEXT)
  var servicesColumn    = getColumnDataByName(sheet, ACCOUNTSSERVICES_CELL_TEXT)
  var filterColumn      = getColumnDataByName(sheet, FILTER_CELL_TEXT)

  var endDate = new Date(startDate. getFullYear(), startDate.getMonth()+1, 0)

  for (var i = 0; i < metricTypeColumn.length; i++) {
    switch (metricTypeColumn[i][0]) {
      case "PagerDuty":
        switch (queryMetricColumn[i][0]) {
          case "Incident Count":
            var incidentCount = getPagerDutyTotalIncidentsByServiceAndTime(
                                  cipMetrics_getPagerdutyAPIToken(),
                                  startDate, endDate,
                                  servicesColumn[i],
                                  filterColumn[i])
            sheet.getRange(i+2,sheet.getActiveCell().getColumn(),1).setValue(incidentCount)
            break;
          case "Incident Duration":
            var incidentDuration = getPagerDutyIncidentsDurationByServiceAndTime(
                                  cipMetrics_getPagerdutyAPIToken(),
                                  startDate, endDate,
                                  servicesColumn[i],
                                  filterColumn[i])
            sheet.getRange(i+2,sheet.getActiveCell().getColumn(),1).setValue(incidentDuration)
            break;
        }
        break;
    }
  }
}
