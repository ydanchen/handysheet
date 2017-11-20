# Simple wrapper for Google Sheets API
 This library provides a layer for easy work with Google Spreadsheets. For example:
 
```
// Create Sheets API client service
Sheets service = SheetServiceProvider.createSheetsService();

// Create a spreadsheet 
SpreadSheet sheet = new SpreadSheet.Builder(service)
       .withId(YOUR_SPREADSHEET_ID_HERE)
       .inRange("Sheet1!A1:A2")
       .build();

// Insert 2 blank rows at the top of sheet       
sheet.insertRows(0, 2, false);        
```
# License
This project is licensed under the terms of the MIT license.