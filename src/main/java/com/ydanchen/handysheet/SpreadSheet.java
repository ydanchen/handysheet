package com.ydanchen.handysheet;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import com.ydanchen.handysheet.enums.Dimension;
import com.ydanchen.handysheet.enums.InputOptionValue;
import com.ydanchen.handysheet.util.Utils;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

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

    /**
     * Constructor
     *
     * @param service an authorized Sheets API client service
     */
    public SpreadSheet(Sheets service) {
        this.service = service;
    }

    // =====================================
    // Setters
    // =====================================

    /**
     * Range setter
     *
     * @param range the new Range. Should match pattern like "A1:E4"
     * @return current instance of the {@link SpreadSheet}
     */
    public SpreadSheet inRange(String range) {
        this.range = range;
        return this;
    }

    /**
     * Sheet setter
     *
     * @param sheet the name of the sheet. Default sheet name is usual "Sheet1"
     * @return current instance of the {@link SpreadSheet}
     */
    public SpreadSheet onSheet(String sheet) {
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

    // =====================================
    // Operations
    // =====================================

    /**
     * Read values from the spreadsheet
     * <p>The range should be specified before with {@code .inRange()} method
     *
     * @return the list of values
     * @throws IOException might be thrown
     */
    public List<List<Object>> getValues() throws IOException {
        return getValuesApiCall();
    }

    /**
     * Read values from the spreadsheet and returns them as two dimensional array
     * <p>The range should be specified before with {@code .inRange()} method
     *
     * @return the readed values
     * @throws IOException might be thrown
     */
    public Object[][] getValuesAsArray() throws IOException {
        return Utils.listOfListsToTwoDimArray(getValuesApiCall());
    }

    /**
     * Update values on the spreadsheet from the List of Lists
     * <p>The range should be specified before with {@code .inRange()} method
     *
     * @param values the values to set
     * @return {@link UpdateValuesResponse}
     * @throws IOException might be thrown
     */
    public UpdateValuesResponse updateValues(List<List<Object>> values) throws IOException {
        return updateValuesApiCall(values);
    }

    /**
     * Update values on the spreadsheet with the values taking from the two dimensional array
     * <p>The range should be specified before with {code}.inRange(){code} method
     *
     * @param values the values to set
     * @return {@link UpdateValuesResponse}
     * @throws IOException might be thrown
     */
    public UpdateValuesResponse updateValues(Object[][] values) throws IOException {
        return updateValuesApiCall(Utils.twoDimArrayToListOfLists(values));
    }

    /**
     * Append values at the end of specified range
     * <p>The range should be specified before with {code}.inRange(){code} method
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
     * <p>The range should be specified before with {code}.inRange(){code} method
     *
     * @param values the values to append
     * @return {@link AppendValuesResponse}
     * @throws IOException might be thrown
     */
    public AppendValuesResponse appendValues(Object[][] values) throws IOException {
        return appendValuesApiCall(Utils.twoDimArrayToListOfLists(values));
    }

    /**
     * Insert new empty rows into the spreadsheet
     *
     * @param startIndex        row index to start from
     * @param endIndex          row index where to end
     * @param inheritFromBefore true to inherit range properties from dimension before,
     *                          false to inherit range properties from dimension after.
     *                          Can't be true if startIndex is 0!
     * @return {@link BatchUpdateSpreadsheetResponse}
     * @throws IOException might be thrown
     */
    public BatchUpdateSpreadsheetResponse insertRows(int startIndex, int endIndex, boolean inheritFromBefore) throws IOException {
        return insertRowsColumnsApiCall(Dimension.ROWS, startIndex, endIndex, inheritFromBefore);
    }

    /**
     * Insert new empty columns into the spreadsheet
     *
     * @param startIndex        row index to start from
     * @param endIndex          row index where to end
     * @param inheritFromBefore true to inherit range properties from dimension before,
     *                          false to inherit range properties from dimension after
     *                          Can't be true if startIndex is 0!
     * @return {@link BatchUpdateSpreadsheetResponse}
     * @throws IOException might be thrown
     */
    public BatchUpdateSpreadsheetResponse insertColumns(int startIndex, int endIndex, boolean inheritFromBefore) throws IOException {
        return insertRowsColumnsApiCall(Dimension.COLUMNS, startIndex, endIndex, inheritFromBefore);
    }

    // =====================================
    // Private methods
    // =====================================

    /**
     * Inserts Rows or Columns to the spreadsheet
     *
     * @param dimension         What we want to insert. Rows or Columns
     * @param startIndex        start index
     * @param endIndex          end index
     * @param inheritFromBefore true to inherit
     * @return {@link BatchUpdateSpreadsheetResponse}
     * @throws IOException will be thrown if occurs
     */
    private BatchUpdateSpreadsheetResponse insertRowsColumnsApiCall(Dimension dimension, int startIndex, int endIndex, boolean inheritFromBefore) throws IOException {
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
     * @param dimension
     * @param startIndex
     * @param endIndex
     * @return
     * @throws IOException
     */
    private BatchUpdateSpreadsheetResponse deleteRowsColumnsApiCall(Dimension dimension, int startIndex, int endIndex) throws IOException {
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