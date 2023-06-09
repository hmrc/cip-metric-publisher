//
// Functions for calling Pager Duty APIs.  
//

//
// Return a date in YYYY-MM-DD format
//
function toYYYYMMDD(day) {
  var d = new Date(day);

  return [
    d.getFullYear(),
    ('0' + (d.getMonth() + 1)).slice(-2),
    ('0' + d.getDate()).slice(-2)
  ].join('-');
}

function test_getPagerDutyIncidentsByServiceAndTime() {
  var startDay = new Date("2023-05-01")
  var endDay = new Date("2023-05-31")
  var incidents = getPagerDutyIncidentsByServiceAndTime(cipMetrics_getPagerdutyAPIToken(), startDay, endDay, "PLHT544")
  var filteredIncidents = filterIncidentsByTitleRegexp(incidents.incidents,"")
  Logger.log(incidents.incidents.length)
  Logger.log(filteredIncidents.length)
}


function filterIncidentsByTitleRegexp (incidents,titleRegexp) {
    var regExp = new RegExp(titleRegexp);
    filteredIncidents = []
    for (var i = 0; i < incidents.length; i++) {
//      Logger.log(incidents[i].title)
      var match = regExp.exec(incidents[i].title);
      if (match) {
//        Logger.log("Match! " + incidents[i].title)
        filteredIncidents.push (incidents[i])
      }
    }
    return filteredIncidents
}


function test_getPagerDutyTotalIncidentsByServiceAndTime() {
  var startDay = new Date("2023-05-01")
  var endDay = new Date("2023-05-31")
  var total = getPagerDutyTotalIncidentsByServiceAndTime(cipMetrics_getPagerdutyAPIToken(), startDay, endDay, "PLHT544","check_cip-data-product-registry-api_exceptions_dynamic_production")
  Logger.log(total)
}

function getPagerDutyTotalIncidentsByServiceAndTime(pgToken,startDay,endDay,serviceIds,titleRegexp) {
  var incidents = getPagerDutyIncidentsByServiceAndTime(pgToken, startDay, endDay, serviceIds).incidents
  var filteredIncidents = filterIncidentsByTitleRegexp(incidents,titleRegexp)
  return filteredIncidents.length
}

function test_getPagerDutyIncidentsDurationByServiceAndTime() {
  var startDay = new Date("2023-05-01")
  var endDay = new Date("2023-05-31")
  var duration = getPagerDutyIncidentsDurationByServiceAndTime(cipMetrics_getPagerdutyAPIToken(), startDay, endDay, "PLHT544","check_cip-data-product-registry-api_exceptions_dynamic_production")
  Logger.log(duration)
}

function getPagerDutyIncidentsDurationByServiceAndTime(pgToken,startDay,endDay,serviceIds,titleRegexp) {
  var incidents = getPagerDutyIncidentsByServiceAndTime(pgToken, startDay, endDay, serviceIds).incidents
  var filteredIncidents = filterIncidentsByTitleRegexp(incidents,titleRegexp)
  totalDuration = 0
  for (var i = 0; i < filteredIncidents.length; i++) {
    Logger.log(filteredIncidents[i].title)
    var created_at = filteredIncidents[i].created_at
    var resolved_at = filteredIncidents[i].resolved_at
//    Logger.log(created_at + " - " + resolved_at)
    if (resolved_at) {
      var diffInSeconds = Math.floor((new Date(resolved_at) - new Date(created_at))/1000);
      totalDuration += diffInSeconds
      Logger.log(diffInSeconds)
    }
  }
  return totalDuration
}

function getPagerDutyIncidentsByServiceAndTime(pgToken,startDay,endDay,serviceIds) {
  var formattedDay = toYYYYMMDD(startDay);
//  var nextDay = day;
//  nextDay.setDate(nextDay.getDate() + 1);
  var formattedNextDay = toYYYYMMDD(endDay);

  var MESSAGE_ENDPOINT = 'https://api.pagerduty.com/incidents?&limit=10000&since=' + formattedDay + '&until=' + formattedNextDay + '&service_ids%5B%5D=' + serviceIds + "";
  var options = {
      method: 'GET',
      muteHttpExceptions: true,
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/vnd.pagerduty+json;version=2',
        'Authorization': 'Token token=' + pgToken
      }
    };
  var res = UrlFetchApp.fetch(MESSAGE_ENDPOINT,options);
  if (res.getResponseCode() != 200) {
    throw new Error(
      "Pagerduty get incidents failed with error code " + res.getResponseCode() + ": " + res
    );
  } else {
//    Logger.log(res.getResponseCode());
  }
  var jsonRes = JSON.parse(res);

  return(jsonRes)
}