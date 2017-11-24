package com.ydanchen.handysheet;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import com.ydanchen.handysheet.enums.Dimension;
import com.ydanchen.handysheet.enums.InputOptionValue;
import com.ydanchen.handysheet.util.Utils;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * This class provides access to the most common Google SpreadSheet operations
 * Built over Google Spreadsheet API v4
 *
 * @author Yevhen Danchenko
 */
public class SpreadSheet {
    private final static String EXC_MARK = "!";
    private Sheets service;
    private String spreadsheetId;
    private String sheet;
    private String range;
    private InputOptionValue inputOptionValue = InputOptionValue.USER_ENTERED;
    private Dimension dimension;
    private int startIndex;
    private int endIndex;

    /**
     * Constructor
     *
     * @param service an authorized Sheets API client service
     */
    public SpreadSheet(Sheets service) {
        this.service = service;
    }

    // =====================================
    // DSL methods
    // =====================================

    /**
     * Range setter
     * <p>This one is for writing semantic
     *
     * @param range the new Range. Should match pattern like "A1:E4"
     * @return current instance of the {@link SpreadSheet}
     */
    public SpreadSheet toRange(String range) {
        this.range = range;
        return this;
    }

    /**
     * Range setter
     * <p>This one is for reading semantic
     *
     * @param range the new Range. Should match pattern like "A1:E4"
     * @return current instance of the {@link SpreadSheet}
     */
    public SpreadSheet fromRange(String range) {
        return toRange(range);
    }

    /**
     * Sheet (tab) setter
     *
     * @param sheet the name of the sheet (tab).
     *              Default sheet name in the new created spreadsheet is usual "Sheet1"
     * @return current instance of the {@link SpreadSheet}
     */
    public SpreadSheet useTab(String sheet) {
        this.sheet = sheet;
        return this;
    }

    /**
     * Spreadsheet ID setter
     *
     * @param spreadsheetId the id of the spreadsheet
     * @return current instance of the {@link SpreadSheet}
     */
    public SpreadSheet withId(String spreadsheetId) {
        this.spreadsheetId = spreadsheetId;
        return this;
    }

    /**
     * Input Value Option setter
     *
     * @param inputValueOption the Input Option Value
     * @return current instance of the {@link SpreadSheet}
     */
    public SpreadSheet withInputValueOption(InputOptionValue inputValueOption) {
        this.inputOptionValue = inputValueOption;
        return this;
    }

    /**
     * Dimension setter. To select Rows or Columns
     *
     * @param dimension the dimension to set
     * @return current instance of the {@link SpreadSheet}
     */
    public SpreadSheet select(Dimension dimension) {
        this.dimension = dimension;
        return this;
    }

    /**
     * Start index setter for dimensions
     *
     * @param startIndex the start index. Can't be lower than 0
     * @return current instance of the {@link SpreadSheet}
     */
    public SpreadSheet from(int startIndex) {
        this.startIndex = startIndex;
        return this;
    }

    /**
     * End index setter for dimensions
     *
     * @param endIndex the end index. Can't be lower than 0
     * @return current instance of the {@link SpreadSheet}
     */
    public SpreadSheet to(int endIndex) {
        this.endIndex = endIndex;
        return this;
    }

    // =====================================
    // Operations
    // =====================================

    /**
     * Read values from the spreadsheet
     * <p>The range should be specified before with {@code .toRange()} method
     *
     * @return the list of values
     * @throws IOException might be thrown
     */
    public List<List<Object>> getValues() throws IOException {
        return getValuesApiCall();
    }

    /**
     * Read values from the spreadsheet and returns them as two dimensional array
     * <p>The range should be specified before with {@code .toRange()} method
     *
     * @return the readed values
     * @throws IOException might be thrown
     */
    public Object[][] getValuesAsArray() throws IOException {
        return Utils.listOfListsToTwoDimArray(getValuesApiCall());
    }

    /**
     * Write values on the spreadsheet from the List of Lists
     * <p>The range should be specified before with {@code .toRange()} or {@code .fromRange()} methods
     *
     * @param values the values to set
     * @return {@link UpdateValuesResponse}
     * @throws IOException might be thrown
     */
    public UpdateValuesResponse writeValues(List<List<Object>> values) throws IOException {
        return updateValuesApiCall(values);
    }

    /**
     * Write values on the spreadsheet with the values taking from the two dimensional array
     * <p>The range should be specified before with {@code .toRange()} or {@code .fromRange()} methods
     *
     * @param values the values to set
     * @return {@link UpdateValuesResponse}
     * @throws IOException might be thrown
     */
    public UpdateValuesResponse writeValues(Object[][] values) throws IOException {
        return updateValuesApiCall(Utils.twoDimArrayToListOfLists(values));
    }

    /**
     * Append values at the end of specified range
     * <p>The range should be specified before with {code}.toRange(){code} method
     *
     * @param values the values to append
     * @return {@link AppendValuesResponse}
     * @throws IOException might be thrown
     */
    public AppendValuesResponse appendValues(List<List<Object>> values) throws IOException {
        return appendValuesApiCall(values);
    }

    /**
     * Append values at the end of specified range
     * <p>The range should be specified before with {code}.toRange(){code} method
     *
     * @param values the values to append
     * @return {@link AppendValuesResponse}
     * @throws IOException might be thrown
     */
    public AppendValuesResponse appendValues(Object[][] values) throws IOException {
        return appendValuesApiCall(Utils.twoDimArrayToListOfLists(values));
    }

    /**
     * Insert new empty rows or columns into the spreadsheet
     *
     * @param inheritFromBefore true to inherit range properties from dimension before,
     *                          false to inherit range properties from dimension after.
     *                          Can't be true if startIndex is 0!
     * @return {@link BatchUpdateSpreadsheetResponse}
     * @throws IOException might be thrown
     */
    public BatchUpdateSpreadsheetResponse insertEmpty(boolean inheritFromBefore) throws IOException {
        return insertRowsColumnsApiCall(inheritFromBefore);
    }

    // =====================================
    // Private methods
    // =====================================

    /**
     * Inserts Rows or Columns to the spreadsheet
     *
     * @param inheritFromBefore true to inherit
     * @return {@link BatchUpdateSpreadsheetResponse}
     * @throws IOException will be thrown if occurs
     */
    private BatchUpdateSpreadsheetResponse insertRowsColumnsApiCall(boolean inheritFromBefore) throws IOException {
        List<Request> requests = new ArrayList<>();
        DimensionRange range = new DimensionRange()
                .setDimension(dimension.getValue())
                .setStartIndex(startIndex)
                .setEndIndex(endIndex);
        requests.add(new Request().setInsertDimension(new InsertDimensionRequest()
                .setInheritFromBefore(inheritFromBefore)
                .setRange(range)));
        BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest();
        requestBody.setRequests(requests);
        Sheets.Spreadsheets.BatchUpdate request = service.spreadsheets().batchUpdate(spreadsheetId, requestBody);
        return request.execute();
    }

    /**
     * Delete Rows and Columns
     * TODO: complete this
     *
     * @return {@link BatchUpdateSpreadsheetRequest}
     * @throws IOException
     */
    private BatchUpdateSpreadsheetResponse deleteRowsColumnsApiCall() throws IOException {
        List<Request> requests = new ArrayList<>();
        DimensionRange range = new DimensionRange()
                .setDimension(dimension.getValue())
                .setStartIndex(startIndex)
                .setEndIndex(endIndex);
        DeleteDimensionRequest deleteDimensionRequest = new DeleteDimensionRequest();
        deleteDimensionRequest.setRange(range);
        BatchUpdateSpreadsheetRequest content = new BatchUpdateSpreadsheetRequest();
        return null;
    }

    /**
     * Update values on the sheet
     *
     * @param values the values to write
     * @return {@link UpdateValuesResponse}
     * @throws IOException will be thrown if occurs
     */
    private UpdateValuesResponse updateValuesApiCall(List<List<Object>> values) throws IOException {
        ValueRange body = new ValueRange().setValues(values);
        return service.spreadsheets().values().update(spreadsheetId, sheet + EXC_MARK + range, body)
                .setValueInputOption(inputOptionValue.getValue())
                .execute();
    }

    /**
     * Append values in the sheet
     *
     * @param values the values to write
     * @return {@link AppendValuesResponse}
     * @throws IOException will be thrown if occurs
     */
    private AppendValuesResponse appendValuesApiCall(List<List<Object>> values) throws IOException {
        ValueRange body = new ValueRange().setValues(values);
        return service.spreadsheets().values().append(spreadsheetId, sheet + EXC_MARK + range, body)
                .setValueInputOption(inputOptionValue.getValue())
                .execute();
    }

    /**
     * Get values from the range specified before
     *
     * @return the list of list of Object
     * @throws IOException will be thrown if occurs
     */
    private List<List<Object>> getValuesApiCall() throws IOException {
        ValueRange response = service.spreadsheets().values()
                .get(spreadsheetId, sheet + EXC_MARK + range)
                .execute();
        return response.getValues();
    }
}