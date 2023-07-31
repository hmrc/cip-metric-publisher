var REPORT_MONTH_TEXT = "Report Month"
var METRIC_ID_CELL_TEXT = "Metric ID"
var METRIC_NAME_CELL_TEXT = "Metric Name"
var RESULTS_CELL_TEXT = "Results"

function configurePublishMenu() {
    fileUrl = PropertiesService.getUserProperties().getProperty('cip-last-slides-url')
    if (fileUrl) {
        SpreadsheetApp.getUi()
          .createMenu('CIP Metrics')
          .addItem('Publish Data to a Slides Pack', 'publishDataHandler')
          .addItem('Publish to ' + PropertiesService.getUserProperties().getProperty('cip-last-slides-name'), 'publishLastFileHandler')
          .addItem('Mark Slide document properties yellow', 'cipMetrics_markSlideDocumentElementsYellow')
          .addSeparator()
          .addItem('Query Metrics for Month', 'cipMetrics_queryMetricsForColumn')
          .addSeparator()
          .addItem('Configure Metric Properties', 'cipMetrics_setupProperties')
          .addToUi();
    } else {
        SpreadsheetApp.getUi()
          .createMenu('CIP Metrics')
          .addItem('Publish Data to a Slides Pack', 'publishDataHandler')
          .addItem('Mark Slide document properties yellow', 'cipMetrics_markSlideDocumentElementsYellow')
          .addSeparator()
          .addItem('Query Metrics for Month', 'cipMetrics_queryMetricsForColumn')
          .addSeparator()
          .addItem('Configure Metric Properties', 'cipMetrics_setupProperties')
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
    publishDataToGoogleSlideFile(fileUrl, configuration);

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
    publishDataToGoogleSlideFile(fileUrl, configuration);
}

//
// This works out the month we're publishing metrics for and creates a data lookup table, returning both
//
function deriveMetricsConfiguration() {
  sheet = SpreadsheetApp.getActiveSheet()
  var metricDate, metricParseType, valuesColumn;

  var reportMonthCellLocation = sheet.createTextFinder(REPORT_MONTH_TEXT).matchEntireCell(true).findNext();

  if(!reportMonthCellLocation) { // if no 'Report Month' found, look at first row and get the active column and use the date from that
      metricDate = new Date(sheet.getRange(1,sheet.getActiveCell().getColumn(),1).getValues()[0][0]);
      metricParseType = 'monthBasedList'
  } else { //grab the date from the cell adjacent to 'Report Month'
      metricDate = new Date(sheet.getRange(reportMonthCellLocation.getRow(),reportMonthCellLocation.getColumn()+1).getValues()[0][0]);
      metricParseType = 'dateField'
  }

  if (isNaN(metricDate.getFullYear())) {
      throw "Couldn't find a valid report publish date to use"
  }

  var metricIdLocation = sheet.createTextFinder(METRIC_ID_CELL_TEXT).matchEntireCell(true).findNext();
  if (!metricIdLocation)  {
      throw "Couldn't find " + METRIC_ID_CELL_TEXT + " marker, did you delete a cell called '" + METRIC_ID_CELL_TEXT + "'?"
  }
  var metricNameLocation = sheet.createTextFinder(METRIC_NAME_CELL_TEXT).matchEntireCell(true).findNext();
  if (metricParseType==='monthBasedList' && !metricNameLocation)  {
      throw "Couldn't find " + METRIC_NAME_CELL_TEXT + " marker, did you delete a cell called '" + METRIC_NAME_CELL_TEXT + "'?"
  }
  var keysColumn   = sheet.getRange(metricIdLocation.getRow()+1,metricIdLocation.getColumn(),sheet.getLastRow()-metricIdLocation.getRow()).getDisplayValues()

  switch (metricParseType) {
    case "monthBasedList":
        valuesColumn   = sheet.getRange(metricIdLocation.getRow()+1,sheet.getActiveCell().getColumn(),keysColumn.length).getDisplayValues()
        var namesColumn    = sheet.getRange(metricNameLocation.getRow()+1,
                                            metricNameLocation.getColumn(),
                                            keysColumn.length,
                                            1).getValues()
        var historyData = sheet.getRange(metricIdLocation.getRow()+1,
                                           1,
                                           keysColumn.length,
                                           sheet.getActiveCell().getColumn()).getValues()
        var titles = sheet.getRange(metricIdLocation.getRow(),
                                    1,
                                    1,
                                    sheet.getActiveCell().getColumn()).getDisplayValues()[0]
//        Logger.log(historyData[0])
        break;
    case "dateField":
        var resultsLocation = sheet.getRange(metricIdLocation.getRow(),metricIdLocation.getColumn(),1,sheet.getLastRow()).createTextFinder(RESULTS_CELL_TEXT).matchEntireCell(true).findNext();
        if (!resultsLocation)  {
            throw "Couldn't find " + RESULTS_CELL_TEXT + " marker, did you delete a cell called '" + RESULTS_CELL_TEXT + "'?"
        }
        valuesColumn   = sheet.getRange(resultsLocation.getRow()+1,resultsLocation.getColumn(),keysColumn.length).getDisplayValues()

        break;
    default:
        throw "Unknown parsing type " + metricParseType
  }

  valueMap   = {}
  nameMap   = {}
  historyMap = {}
  for (var i = 0; i < keysColumn.length; i++) {
      if (keysColumn[i][0] && keysColumn[i][0] != "") {
//        Logger.log(keysColumn[i][0] + ' -> ' + valuesColumn[i][0])
        valueMap[keysColumn[i][0]] = valuesColumn[i][0]
        if (metricParseType==='monthBasedList') {
            nameMap[keysColumn[i][0]] = namesColumn[i][0]
        }
        if (historyData) {
          historyMap[keysColumn[i][0]] = historyData[i]
        }
      }
  }

  var resultMap = {}
  resultMap["month"] = metricDate.getFullYear() + '-' + (metricDate.getMonth()+1)
  resultMap["valueMap"] = valueMap
  resultMap["nameMap"] = nameMap
  resultMap["historyMap"] = historyMap
  if (titles) resultMap["titles"] = titles
  return resultMap
}


// Loop through each page element, and if it has a source marker, and we have value for it,
// set the value of the element.
function publishDataToGoogleSlideFile_processSlide(slide, configuration, tmpSlidesDoc) {
//    Logger.log(slide)
    var valueMap = configuration["valueMap"]
    var elements = slide.getPageElements()

    // Loop through elements identifying if any have Source references in the alt text
    for (var i = 0; i < elements.length; i++) {
        var element = elements[i]
        var alt = element.getDescription()

        var regExp = new RegExp('.*(Source|RotateImage|Graph|Render): ([^\\s]+)', "i");

        var match = regExp.exec(alt);
        Logger.log("Deciding on " + alt + " -------------------- " + match)
        if (match && match.length===3 && match[1]==='Source') {
            var sourceId = match[2]
            // If we've got a value for the Source, set the page elements value
            if (valueMap[sourceId]) {
                var regExp2 = new RegExp('.*-percentage-change', "i");
                var matchPercentageChange = regExp2.exec(sourceId);
                var finalVal = valueMap[sourceId];
                Logger.log("Processing value " + finalVal)
                if (matchPercentageChange) {
                    Logger.log("Found a percentage change to construct");
                    finalVal = '('+(finalVal.replace("%",""))+'%)'
                }
                Logger.log("Setting " + sourceId + " to " + finalVal)
                element.asShape().getText().setText(finalVal)
                element.asShape().getText().getTextStyle().setBackgroundColorTransparent()
            }
        } else if (match && match.length===3 && match[1]==='RotateImage') {
	    Logger.log("Processing rotate image: " + alt)
            var sourceId = match[2]
            if (valueMap[sourceId]) {
                var rawPercentageChange = valueMap[sourceId].replace("%","");

                if (rawPercentageChange < 0) {
                    Logger.log("Rotating 180 degrees");
                    element.asShape().setRotation(180);
                } else if (rawPercentageChange > 0) {
                    Logger.log("Rotating 0 degrees");
                    element.asShape().setRotation(0);
                }  else if (rawPercentageChange == 0) {
                    Logger.log("Rotating 270 degrees");
                    element.asShape().setRotation(270);
                }
            }
        } else if (match && match.length===3 && match[1]==='Graph') {
            var sourceId = match[2]
	          if (valueMap[sourceId] && element.getPageElementType() == "IMAGE") {
		        // Slightly unpleasant code to convert a graph to an image via a temporary slide document.
		        // This is beacuse the getAs method on the Gooogle Sheet graph object is broken, and returns axis with incorrect
		        // labels.

		        foundSourceGraph = searchGraphByName(valueMap[sourceId])
		        if (foundSourceGraph) {
              var imageBlob = tmpSlidesDoc.getSlides()[0].insertSheetsChartAsImage(foundSourceGraph).getAs("image/png");
              element.asImage().replace(imageBlob)
		          element.setDescription(alt)
		        }
	        }
        } else if (match && match.length===3 && match[1]==='Render') {
          var sourceIds = match[2].split(",")
          Logger.log("Render " + sourceIds)
	        if (element.getPageElementType() == "IMAGE") {
            graphConfig = "{}"
            var renderRegExp = new RegExp('.*(Render): ([^\\s]+) (.*)$', "i");
            Logger.log("Matching against :" + alt)
            var renderMatch = renderRegExp.exec(alt);
            Logger.log("RenderMatch = " + renderMatch.length)
            if (renderMatch.length === 4) {
              graphConfig = renderMatch[3]
            }
            var imageBlob = createGraphBlob(configuration, sourceIds, graphConfig)
            element.asImage().replace(imageBlob)
		        element.setDescription(alt)
          }

        } else {
          Logger.log("Couldn't match " + alt)
        }
   } // elements loop
}


function publishDataToGoogleSlideFile(documentUrl, configuration) {
    var slideDoc = SlidesApp.openByUrl(documentUrl)
    var slides = slideDoc.getSlides()
    var tmpSlidesDoc = SlidesApp.create("temp_slide_for_image")
    for (var i = 0; i < slides.length; i++) {
        publishDataToGoogleSlideFile_processSlide(slides[i], configuration, tmpSlidesDoc)
    }
    DriveApp.getFileById(tmpSlidesDoc.getId()).setTrashed(true);
}


function cipMetrics_markSlideDocumentElementsYellow() {

  var ui = SpreadsheetApp.getUi();

  var slideUrlHandler = ui.prompt(
      'Target slides file',
      'Please enter URL for the target slide document to mark yellow',
      ui.ButtonSet.OK_CANCEL);

  documentUrl = slideUrlHandler.getResponseText();

  if (documentUrl) {
    var slideDoc = SlidesApp.openByUrl(documentUrl)
    var slides = slideDoc.getSlides()
    for (var i = 0; i < slides.length; i++) {
      var elements = slides[i].getPageElements()
      // Loop through elements identifying if any have Source references in the alt text
      for (var e = 0; e < elements.length; e++) {
        var element = elements[e]
        var alt = element.getDescription()
        if (alt != "") {
          try {
            var textStyle = element.asShape().getText().getTextStyle().setBackgroundColor("#f0fc27")
          } catch (err) {
            Logger.log(err)
          }
        }
      }
    }
  }
}
