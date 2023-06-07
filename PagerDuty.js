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
  var incidentCount = getPagerDutyIncidentsByServiceAndTime(cipMetrics_getPagerdutyAPIToken(), startDay, endDay, "PLHT544")
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

  return(jsonRes.incidents)
}