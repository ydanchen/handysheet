# Simple Java DSL to work with Google Sheets API
 This library provides a layer for easy work with Google Spreadsheets API v4. For example:
 
```
// Create Sheets API client service
Sheets service = SheetsServiceProvider.createSheetsService(APPLICATION_NAME);

// Create a spreadsheet 
SpreadSheet sheet = new SpreadSheet(service).withId(SPREEDSHEET_ID);

// Write values
sheet.onTab("Sheet1").toRange("A1:C3").writeValues(values);

// Insert one blank row at the top of a sheet       
sheet.onTab("Sheet1").select(Dimension.ROWS).from(0).to(1).insertEmpty();        
```
# License
This project is licensed under the terms of the MIT license.