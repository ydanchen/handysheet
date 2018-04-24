# Simple Java DSL to work with Google Sheets API
 This library provides a layer for easy work with Google Spreadsheets API v4. For example:
 
```java
// Create Sheets API client service
Sheets service = SheetsServiceProvider.createSheetsService(APPLICATION_NAME);

// Connect to a spreadsheet 
SpreadSheet spreadsheet = new SpreadSheet(service).withId(SPREEDSHEET_ID);

// Write values
spreadsheet.onSheet("Sheet1").toRange("A1:C3").writeValues(values);

// Insert one blank row at the top of a sheet       
spreadsheet.onSheet("Sheet1").select(Dimension.ROWS).from(0).to(1).insertEmpty();        
```

# More examples
Append values:
```java
spreadsheet
       .onSheet("Sheet1")
       .toRange("A4:E4")
       .withValueInputOption(ValueInputOption.RAW)
       .appendValues(values);
```
Insert empty rows or columns:
```java
spreadsheet
       .onSheet("Sheet1")
       .select(Dimension.ROWS)
       .from(0)
       .to(1)
       .insertEmpty();
```
Merge cells:
```java
spreadsheet
       .onSheet("Sheet1")
       .from(2,2)
       .to(4,4)
       .mergeCells();
```

# License
This project is licensed under the terms of the MIT license.
