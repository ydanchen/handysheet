package com.ydanchen.handysheet;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import com.ydanchen.handysheet.enums.Dimension;
import com.ydanchen.handysheet.enums.MergeType;
import com.ydanchen.handysheet.enums.SortOrder;
import com.ydanchen.handysheet.enums.ValueInputOption;
import com.ydanchen.handysheet.util.Utils;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

/**
 * This class provides access to the most common Google SpreadSheet operations
 * Built over the Google Spreadsheet API v4
 *
 * @author Yevhen Danchenko
 */
public class SpreadSheet {
    private final static String EXCLAMATION_MARK = "!";

    private Sheets service;
    private String spreadsheetId;
    private String sheet;
    private String range;
    private ValueInputOption valueInputOption = ValueInputOption.USER_ENTERED;
    private Boolean inheritFromBefore = false;
    private Dimension dimension;
    private MergeType mergeType = MergeType.MERGE_ALL;
    private SortOrder sortOrder = SortOrder.ASCENDING;
    private int startIndex;
    private int endIndex;
    private int startRowIndex;
    private int startColumnIndex;
    private int endRowIndex;
    private int endColumnIndex;

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
     * Sheet setter
     *
     * @param sheet the name of the sheet.
     *              Default sheet name in the new created spreadsheet is usual "Sheet1"
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
     * Value Input Option setter
     *
     * @param valueInputValue the Value Input Option {@link ValueInputOption}
     * @return current instance of the {@link SpreadSheet}
     */
    public SpreadSheet withValueInputOption(ValueInputOption valueInputValue) {
        this.valueInputOption = valueInputValue;
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
     * Start indexes for methods when complete range is required, e.g. mergeCells()
     *
     * @param startColumnIndex the column start index. Can't be lower than 0
     * @param startRowIndex    the row start index. Can't be lower than 0
     * @return current instance of the {@link SpreadSheet}
     */
    public SpreadSheet from(int startColumnIndex, int startRowIndex) {
        this.startColumnIndex = startColumnIndex;
        this.startRowIndex = startRowIndex;
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

    /**
     * End indexes for methods when complete range is required, e.g. mergeCells()
     *
     * @param endColumnIndex the column end index. Can't be lower than 0
     * @param endRowIndex    the row end index. Can't be lower than 0
     * @return current instance of the {@link SpreadSheet}
     */
    public SpreadSheet to(int endColumnIndex, int endRowIndex) {
        this.endColumnIndex = endColumnIndex;
        this.endRowIndex = endRowIndex;
        return this;
    }

    /**
     * Inherit range properties form dimension before setter
     *
     * @param inheritFromBefore true to inherit range properties from dimension before,
     *                          false to inherit range properties from dimension after.
     *                          Can't be true if startIndex is 0!
     * @return current instance of the {@link SpreadSheet}
     */
    public SpreadSheet inheritFromBefore(Boolean inheritFromBefore) {
        this.inheritFromBefore = inheritFromBefore;
        return this;
    }

    /**
     * Merge type setter.
     *
     * @param mergeType specifies how cells should be merged.
     * @return current instance of the {@link SpreadSheet}
     */
    public SpreadSheet mergeBy(MergeType mergeType) {
        this.mergeType = mergeType;
        return this;
    }

    /**
     * Sort order setter
     *
     * @param sortOrder the order data should be sorted.
     * @return current instance of the {@link SpreadSheet}
     */
    public SpreadSheet byOrder(SortOrder sortOrder) {
        this.sortOrder = sortOrder;
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
     * @return {@link BatchUpdateSpreadsheetResponse}
     * @throws IOException might be thrown
     */
    public BatchUpdateSpreadsheetResponse insertEmpty() throws IOException {
        return insertRowsColumnsApiCall();
    }

    /**
     * Delete rows or cloumns
     *
     * @return {@link BatchUpdateSpreadsheetResponse}
     * @throws IOException might be thrown
     */
    public BatchUpdateSpreadsheetResponse delete() throws IOException {
        return deleteRowsColumnsApiCall();
    }

    /**
     * Sort the values on the sheet
     *
     * @return {@link BatchUpdateSpreadsheetResponse}
     * @throws IOException might be thrown
     */
    public BatchUpdateSpreadsheetResponse sort() throws IOException {
        return sortApiCall();
    }

    // =====================================
    // Accessors
    // =====================================

    /**
     * Return all Sheets in the spreadsheet
     *
     * @return list of all Sheets in the spreadsheet
     * @throws IOException might be thrown
     */
    public List<Sheet> getSheets() throws IOException {
        return service.spreadsheets().get(spreadsheetId).execute().getSheets();
    }

    /**
     * Return names of all sheets in the spreadsheet as List of String
     *
     * @return all sheet names as a List of String
     * @throws IOException might be thrown
     */
    public List<String> getSheetsNames() throws IOException {
        return getSheets()
                .stream()
                .map(v -> v.getProperties().getTitle())
                .collect(Collectors.toList());
    }

    // =====================================
    // Formatting
    // =====================================

    /**
     * Merge cells on sheet
     *
     * @return {@link BatchUpdateSpreadsheetResponse}
     * @throws IOException might be thrown
     */
    public BatchUpdateSpreadsheetResponse mergeCells() throws IOException {
        return MergeCellsApiCall();
    }


    // =====================================
    // Private methods (API calls)
    // =====================================

    /**
     * Inserts Rows or Columns to the spreadsheet
     *
     * @return {@link BatchUpdateSpreadsheetResponse}
     * @throws IOException will be thrown if occurs
     */
    private BatchUpdateSpreadsheetResponse insertRowsColumnsApiCall() throws IOException {
        List<Request> requests = new ArrayList<>();
        DimensionRange range = new DimensionRange()
                .setDimension(dimension.getValue())
                .setStartIndex(startIndex)
                .setEndIndex(endIndex);
        requests.add(new Request().setInsertDimension(new InsertDimensionRequest()
                .setInheritFromBefore(inheritFromBefore)
                .setRange(range)));
        BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
        return service.spreadsheets().batchUpdate(spreadsheetId, requestBody).execute();
    }

    /**
     * Delete Rows and Columns on the spreadsheet
     *
     * @return {@link BatchUpdateSpreadsheetRequest}
     * @throws IOException will be thrown if occurs
     */
    private BatchUpdateSpreadsheetResponse deleteRowsColumnsApiCall() throws IOException {
        List<Request> requests = new ArrayList<>();
        DimensionRange range = new DimensionRange()
                .setDimension(dimension.getValue())
                .setStartIndex(startIndex)
                .setEndIndex(endIndex);
        requests.add(new Request().setDeleteDimension(new DeleteDimensionRequest().setRange(range)));
        BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
        return service.spreadsheets().batchUpdate(spreadsheetId, requestBody).execute();
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
        return service.spreadsheets().values().update(spreadsheetId, getRangeWithSheet(sheet, range), body)
                .setValueInputOption(valueInputOption.getValue())
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
        return service.spreadsheets().values().append(spreadsheetId, getRangeWithSheet(sheet, range), body)
                .setValueInputOption(valueInputOption.getValue())
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
                .get(spreadsheetId, getRangeWithSheet(sheet, range))
                .execute();
        return response.getValues();
    }

    /**
     * Merge cells on the grid
     *
     * @return {@link BatchUpdateSpreadsheetResponse}
     * @throws IOException will be thrown if occurs
     */
    private BatchUpdateSpreadsheetResponse MergeCellsApiCall() throws IOException {
        List<Request> requests = new ArrayList<>();
        GridRange range = new GridRange()
                .setStartColumnIndex(startColumnIndex)
                .setStartRowIndex(startRowIndex)
                .setEndColumnIndex(endColumnIndex)
                .setEndRowIndex(endRowIndex);
        requests.add(new Request().setMergeCells(new MergeCellsRequest()
                .setMergeType(mergeType.getValue())
                .setRange(range)));
        BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
        return service.spreadsheets().batchUpdate(spreadsheetId, requestBody).execute();
    }

    /**
     * Sort range on the spreadsheet
     *
     * @return {@link BatchUpdateSpreadsheetRequest}
     * @throws IOException will be thrown if occurs
     */
    private BatchUpdateSpreadsheetResponse sortApiCall() throws IOException {
        List<Request> requests = new ArrayList<>();
        List<SortSpec> sortSpecs = new ArrayList<>();
        sortSpecs.add(
                new SortSpec()
                .setSortOrder(sortOrder.getValue())
                .setDimensionIndex((dimension.equals(Dimension.COLUMNS) ? 0 : 1))
        );
        GridRange range = new GridRange()
                .setStartColumnIndex(startColumnIndex)
                .setStartRowIndex(startRowIndex)
                .setEndColumnIndex(endColumnIndex)
                .setEndRowIndex(endRowIndex);
        requests.add(new Request().setSortRange(new SortRangeRequest()
                .setRange(range)
                .setSortSpecs(sortSpecs)));
        BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);
        return service.spreadsheets().batchUpdate(spreadsheetId, requestBody).execute();
    }

    /**
     * Concatenates sheet name and range to provide range name suitable for API, like "Sheet1!A1:B2"
     *
     * @param sheet the sheet name. The name of default sheet is "Sheet1"
     * @param range the range, e.g. "A1:B2"
     * @return the range suitable for Google Sheets API
     */
    private static String getRangeWithSheet(String sheet, String range) {
        return sheet + EXCLAMATION_MARK + range;
    }
}