//
// This file contains functions to manage sensitive tokens within the users properties rather
// than in locations viewable by all users of the spreadsheet
//
var aesKey = "J@NcRfUjXn2r5u7x!A%D*G-KaPdSgVkY"

function cipMetrics_setupProperties() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var pgTokenHandler = ui.prompt(
      'Setting up PagerDuty API key (step 1 of 1)',
      'Please enter your PagerDuty API token (leave blank if unused):',
      ui.ButtonSet.OK_CANCEL);
  var userProperties = PropertiesService.getUserProperties();
  var cipher = new Cipher(aesKey, 'aes');
  userProperties.setProperties({
    "CIP Metrics Pager Duty API Token" : cipher.encrypt(pgTokenHandler.getResponseText())
  })
}

function cipMetrics_getPagerdutyAPIToken() {
  var userProperties = PropertiesService.getUserProperties();
  var cipher = new Cipher(aesKey, 'aes');

  token = cipher.decrypt(userProperties.getProperty("CIP Metrics Pager Duty API Token"));
  return token;
}

function cipMetrics_getProperties() {
   var userProperties = PropertiesService.getUserProperties();
   var cipher = new Cipher(aesKey, 'aes');
  Logger.log(userProperties.getProperties());
  Logger.log(cipher.decrypt(userProperties.getProperty("CIP Metrics Pager Duty API Token")));
}