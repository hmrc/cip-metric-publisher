
# cip-metric-publisher

This is a placeholder README.md for a new repository

cip-metric-publisher is used for publishing data from a Goggle Sheet to a Google Slide document by use of markers in the Slide documnt to indicate the source of data.

Two formats of Google sheet are supported, one to support a list of keys and values by month, and a second to support Google Analytics reports.

## The Google Sheet Library Integration
To include this library, open the document, then click Extensions -> Apps Script 


## The Google Sheet Format
A Google Sheet represents the source of data, and is in one of two formats

### Key / Month matrix

| Metric ID  | ... | ... | month1 date | month2 date | month3 date |
| :---	     | :-- | :-- |             |             |             |
| unique-id1 |     |     | month1 val  | month2 val  | month3 val  |
| unique-id2 |     |     | month1 val  | month2 val  | month3 val  |

The column "Metric ID" is the marker for the column of unique metric identifiers. Columns titles with values should contain a date, where the date is any date in the month to which the metric value is attirbuted. Testing has been conducted using a 1st day of month, but any date in the month should suffice.

To publish data for a particular month, select any row in the month you wish to publish, and select the Publish to a Slide option under the 'CIP Publish' menu. You'll be prompted for the URL of a destination Google Slide document. Enter the URL and click Ok. Your data will be published to marked up elements (see The Slide Document section for how this is done).

### Google Analytics Config format
For Google Analytics report configuration format, we expect a format as follows:

|              | Report Month | month date |             |             |             |
| :---	       | :--          | :--        |             |             |             |
| Report Name  | Report 1     | Report 2   | Report 3    | Report 4    | Report 5    |
| Config 1     | setting 1    | setting 2  | setting 3   | setting 4   | setting 5   |
| Config 2     | setting 1    | setting 2  | setting 3   | setting 4   | setting 5   |
| Config 3     | setting 1    | setting 2  | setting 3   | setting 4   | setting 5   |
| Config 4     | setting 1    | setting 2  | setting 3   | setting 4   | setting 5   |
|              |              |            |             |             |             |
| Metric ID    | Results      |            |             |             |             |
| unique_id1   | value1       |            |             |             |             |
| unique_id2   | value2       |            |             |             |             |
| unique_id3   | value3       |            |             |             |             |
| unique_id4   | value4       |            |             |             |             |

In this format, the sheet reflects a single month of data. To publish the data, you can be anywhere in the sheet. Select the publish menu option as per Key/Month instructions above. The metrics will be published to the destination document.

## The Slide Document
Page elements are marked to receive data by adding tag information to the Description field of the Alt Text information. This tagging specified the source metric identifier used to populate it. During the process of publishing, slide page elements are searched; any element with one of these tags where the source data has a matching metric identifier, it's data is replaced.

Supported tagging formats are:
| Tag format   | Behaviour    |
| :---	       | :--          |
| Source: <metricId> | Replaces the textual content of the page element with that of the google sheet source. This uses the display format from the google sheet, so any google sheet cell rendering is honoured |
| RotateImage: <metricId> | Used on images. Rotates the image if the metricId has a value of less than zero. Typically used on arrows |

### What is appsscript.json file?

The appsscript.json file is a hidden manifest file that allows the Timezone that the script is running in to be set. For the mdtp-team-scheduler, this should always be set to "Europe/London".

To see this file in the Google AppsScript console, enable it in Project Settings > Show "appsscript.json" manifest file in editor.

## Development

### Requirements
- Node v12.20.1

### Install
Install dependencies using: 
```shell
npm ci
```

## Deployment
To enable Google Apps Script API for your user, go to [Apps Script Settings](https://script.google.com/home/usersettings) and toggle Google Apps Script API to on.

Clasp needs access to push using your Google account. This can be achieved by running:

```shell
npx clasp login
```

This should open a browser which asks you to give Clasp access to the Google App Scripts API.

To deploy changes, run:

```shell
npx clasp push
```

### License

This code is open source software licensed under the [Apache 2.0 License]("http://www.apache.org/licenses/LICENSE-2.0.html").
