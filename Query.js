var METRIC_ID_CELL_TEXT = "Metric ID"
var ACCOUNTSSERVICES_CELL_TEXT = "Accounts/Services"
var METRIC_TYPE_CELL_TEXT = "Metric Type"

function cipMetrics_queryMetricsForColumn() {
  sheet = SpreadsheetApp.getActiveSheet()
  
  startDate = new Date(sheet.getRange(1,sheet.getActiveCell().getColumn(),1).getValues()[0][0]);
  Logger.log(startDate)
  if (isNaN(startDate.getFullYear())) {
      throw "Couldn't find a valid report publish date to use"
  }

  var metricTypeLocation = sheet.createTextFinder(METRIC_TYPE_CELL_TEXT).matchEntireCell(true).findNext();
  if (!metricTypeLocation)  {
      throw "Couldn't find " + METRIC_TYPE_CELL_TEXT + " marker, did you delete a cell called '" + METRIC_TYPE_CELL_TEXT + "'?"
  }
  var metricTypeColumn  = sheet.getRange(metricTypeLocation.getRow()+1,metricTypeLocation.getColumn(),sheet.getLastRow()-metricTypeLocation.getRow()).getDisplayValues()

  var metricIdLocation = sheet.createTextFinder(METRIC_ID_CELL_TEXT).matchEntireCell(true).findNext();
  if (!metricIdLocation)  {
      throw "Couldn't find " + METRIC_ID_CELL_TEXT + " marker, did you delete a cell called '" + METRIC_ID_CELL_TEXT + "'?"
  }
  var keysColumn   = sheet.getRange(metricIdLocation.getRow()+1,metricIdLocation.getColumn(),sheet.getLastRow()-metricIdLocation.getRow()).getDisplayValues()

  var servicesLocation = sheet.createTextFinder(ACCOUNTSSERVICES_CELL_TEXT).matchEntireCell(true).findNext();
  if (!servicesLocation)  {
      throw "Couldn't find " + ACCOUNTSSERVICES_CELL_TEXT + " marker, did you delete a cell called '" + ACCOUNTSSERVICES_CELL_TEXT + "'?"
  }
  var servicesColumn   = sheet.getRange(servicesLocation.getRow()+1,servicesLocation.getColumn(),sheet.getLastRow()-servicesLocation.getRow()).getDisplayValues()
  
  var endDate = new Date(startDate. getFullYear(), startDate.getMonth()+1, 0)

  for (var i = 0; i < 2; i++) { //metricTypeColumn.length; i++) {
    switch (metricTypeColumn[i][0]) {
      case "PG Incidents":
        var incidentCount = getPagerDutyIncidentsByServiceAndTime(
                              cipMetrics_getPagerdutyAPIToken(),
                              startDate, endDate,
                              servicesColumn[i]).length
        sheet.getRange(i+2,sheet.getActiveCell().getColumn(),1).setValue(incidentCount)
    }
  }
}
