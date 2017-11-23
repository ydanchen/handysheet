# Simple DSL to work with Google Sheets API
 This library provides a layer for easy work with Google Spreadsheets. For example:
 
```
// Create Sheets API client service
Sheets service = SheetsServiceProvider.createSheetsService(APPLICATION_NAME);

// Prepare some values 
Object[][] values = {
    {"A1", "B1", "C1"},
    {"A2", "B2", "C2"},
    {"A3", "B3", "C3"}
};

// Create a spreadsheet 
SpreadSheet sheet = new SpreadSheet.Builder(service)
       .withId(YOUR_SPREADSHEET_ID_HERE)
       .inRange("Sheet1!A1:C3")
       .build();

// Write values
sheet.updateValues(values);

// Insert one blank row at the top of a sheet       
sheet.insertRows(0, 1, false);        
```
# License
This project is licensed under the terms of the MIT license.