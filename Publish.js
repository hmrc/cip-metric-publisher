var REPORT_MONTH_TEXT = "Report Month"
var METRIC_ID_CELL_TEXT = "Metric ID"
var RESULTS_CELL_TEXT = "Results"

/* function onOpen() {
   configurePublishMenu()
}
*/

function configurePublishMenu() {
    fileUrl = PropertiesService.getUserProperties().getProperty('cip-last-slides-url')
Logger.log("Last file URL = " + fileUrl)
    if (fileUrl) {
        Logger.log('Last slides file url: ' +fileUrl)
        SpreadsheetApp.getUi()
          .createMenu('CIP Publish')
          .addItem('Publish Data to a Slides Pack', 'publishDataHandler')
          .addItem('Publish to ' + PropertiesService.getUserProperties().getProperty('cip-last-slides-name'), 'publishLastFileHandler')
          .addToUi();
    } else {
        SpreadsheetApp.getUi()
          .createMenu('CIP Publish')
          .addItem('Publish Data to a Slides Pack', 'publishDataHandler')
          .addToUi();
    }
}



function publishDataHandler() {

  configuration = deriveMetricsConfiguration()

  var ui = SpreadsheetApp.getUi();

  var slideUrlHandler = ui.prompt(
      'Target slides file',
      'Publishing data for ' + configuration["month"] + '. Please enter URL for the target slide document to update data',
      ui.ButtonSet.OK_CANCEL);
  
  fileUrl = slideUrlHandler.getResponseText();

  if (fileUrl) {
    publishDataToGoogleSlideFile(fileUrl, configuration["valueMap"]);

    var docName = SlidesApp.openByUrl(fileUrl).getName()
    PropertiesService.getUserProperties().setProperties({
      'cip-last-slides-url' : fileUrl,
      'cip-last-slides-name' : docName
    });

    configurePublishMenu();
  }  
}


function publishLastFileHandler() {
    configuration = deriveMetricsConfiguration()
    fileUrl = PropertiesService.getUserProperties().getProperty('cip-last-slides-url')
    publishDataToGoogleSlideFile(fileUrl, configuration["valueMap"]);
}

//
// This works out the month we're publishing metrics for and creates a data lookup table, returning both
//
function deriveMetricsConfiguration() {
  sheet = SpreadsheetApp.getActiveSheet()
  var metricDate, regExp;

  var reportMonthCellLocation = sheet.createTextFinder(REPORT_MONTH_TEXT).matchEntireCell(true).findNext();
  
  if(!reportMonthCellLocation) { // if no 'Report Month' found, look at first row and get the active column and use the date from that
      metricDate = new Date(sheet.getRange(1,sheet.getActiveCell().getColumn(),1).getValues()[0][0]);
      metricParseType = 'monthBasedList'
  } else { //grab the date from the cell adjacent to 'Report Month'
      metricDate = new Date(sheet.getRange(reportMonthCellLocation.getRow(),reportMonthCellLocation.getColumn()+1).getValues()[0][0]);
      metricParseType = 'googleAnalyticsConfig'
  }

  if (isNaN(metricDate.getFullYear())) {
      throw "Couldn't find a valid report publish date to use"
  }

  var metricIdLocation = sheet.createTextFinder(METRIC_ID_CELL_TEXT).matchEntireCell(true).findNext();
  if (!metricIdLocation)  {
      throw "Couldn't find " + METRIC_ID_CELL_TEXT + " marker, did you delete a cell called '" + METRIC_ID_CELL_TEXT + "'?"
  }
  var keysColumn   = sheet.getRange(metricIdLocation.getRow()+1,metricIdLocation.getColumn(),sheet.getLastRow()-metricIdLocation.getRow()).getDisplayValues()

  switch (metricParseType) {
    case "monthBasedList":
        var valuesColumn   = sheet.getRange(metricIdLocation.getRow()+1,sheet.getActiveCell().getColumn(),keysColumn.length).getDisplayValues()
        break;
    case "googleAnalyticsConfig":
        var resultsLocation = sheet.getRange(metricIdLocation.getRow(),metricIdLocation.getColumn(),1,sheet.getLastRow()).createTextFinder(RESULTS_CELL_TEXT).matchEntireCell(true).findNext();
        if (!resultsLocation)  {
            throw "Couldn't find " + RESULTS_CELL_TEXT + " marker, did you delete a cell called '" + RESULTS_CELL_TEXT + "'?"
        }
        var valuesColumn   = sheet.getRange(resultsLocation.getRow()+1,resultsLocation.getColumn(),keysColumn.length).getDisplayValues()

        break;
    default:
        throw "Unknown parsing type " + metricParseType
  }

  valueMap = {}
  for (var i = 0; i < keysColumn.length; i++) {
      if (keysColumn[i][0] && keysColumn[i][0] != "") {
//          Logger.log(keysColumn[i][0] + ' -> ' + valuesColumn[i][0])
          valueMap[keysColumn[i][0]] = valuesColumn[i][0]
      }
  }
  
  var resultMap = {}
  resultMap["month"] = metricDate.getFullYear() + '-' + (metricDate.getMonth()+1)
  resultMap["valueMap"] = valueMap
  Logger.log(resultMap)
  return resultMap
}


// Loop through each page element, and if it has a source marker, and we have value for it,
// set the value of the element.
function publishDataToGoogleSlideFile_processSlide(slide, valueMap) {
//    Logger.log(slide)
    var elements = slide.getPageElements()

    // Loop through elements identifying if any have Source references in the alt text
    for (var i = 0; i < elements.length; i++) {
        element = elements[i]
        alt = element.getDescription()

        var regExp = new RegExp('.*(Source|RotateImage): ([^\\s]+)', "i"); 
        
        var match = regExp.exec(alt);
        
        if (match && match.length===3 && match[1]==='Source') {
            var sourceId = match[2]
            // If we've got a value for the Source, set the page elements value
            if (valueMap[sourceId]) {
                var regExp2 = new RegExp('.*-percentage-change', "i"); 
                var matchPercentageChange = regExp2.exec(sourceId);
                var finalVal = valueMap[sourceId];
                if (matchPercentageChange) {
                    Logger.log("Found a percentage change to construct");
                    finalVal = '('+Math.abs(finalVal)+'%)'
                }
                Logger.log("Setting " + sourceId + " to " + finalVal)
                element.asShape().getText().setText(finalVal)                
            }
        } else if (match && match.length===3 && match[1]==='RotateImage') {
            Logger.log("Found a shape to rotate...")
            if (valueMap[match[2]] < 0) {
              Logger.log("Rotating 180 degrees");
              element.asShape().setRotation(180);
            } else {
               Logger.log("Rotating 0 degrees");
              element.asShape().setRotation(0);
            } 
        }
    }
}


function publishDataToGoogleSlideFile(documentUrl, valueMap) {
    var slideDoc = SlidesApp.openByUrl(documentUrl)
    var slides = slideDoc.getSlides()
    for (var i = 0; i < slides.length; i++) {
        publishDataToGoogleSlideFile_processSlide(slides[i], valueMap)
    }
}
