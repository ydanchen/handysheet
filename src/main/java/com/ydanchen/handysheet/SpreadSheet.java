package com.ydanchen.handysheet;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import com.ydanchen.handysheet.enums.Dimension;
import com.ydanchen.handysheet.enums.InputOptionValue;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * This class provides access to the most common Google SpreadSheet operations
 *
 * @author Yevhen Danchenko
 */
public class SpreadSheet {
    private Sheets service;
    private String spreadsheetId;
    private String range;

    /**
     * Constructor
     *
     * @param builder com.ydanchen.handysheets.apps.spreadsheet.SpreadSheet Builder
     */
    private SpreadSheet(Builder builder) {
        this.service = builder.service;
        this.spreadsheetId = builder.spreadsheetId;
        this.range = builder.range;
    }

    /**
     * Range setter
     *
     * @param range the new Range
     */
    public void setRange(final String range) {
        this.range = range;
    }

    /**
     * Read values from the specified spreadsheet
     *
     * @return the list of values
     * @throws IOException might be thrown
     */
    public List<List<Object>> getValues() throws IOException {
        ValueRange response = service.spreadsheets().values()
                .get(spreadsheetId, range)
                .execute();
        return response.getValues();
    }

    /**
     * Read values from the specified spreadsheet and returns them as two dimensional array
     *
     * @return the readed values
     * @throws IOException might be thrown
     */
    public Object[][] getValuesAsArray() throws IOException {

        return null;
    }

    /**
     * Set values to the spreadsheet from the List of Lists
     *
     * @param values the values to set
     * @throws IOException might be thrown
     */
    public void setValues(final List<List<Object>> values) throws IOException {
        ValueRange body = new ValueRange().setValues(values);
        UpdateValuesResponse response = service.spreadsheets().values().update(spreadsheetId, range, body)
                .setValueInputOption(InputOptionValue.USER_ENTERED.getValue())
                .execute();
    }

    /**
     * Set values to the spreadsheet from the two dimensional array
     *
     * @param values the values to set
     * @throws IOException might be thrown
     */
    public void setValues(Object[][] values) throws IOException {
        //TODO: TBD
    }

    public void insertRows(int startIndex, int endIndex, boolean inheritFromBefore) throws IOException {
        insertRowsColumns(Dimension.ROWS.getValue(), startIndex, endIndex, inheritFromBefore);
    }

    public void insertColumns(int startIndex, int endIndex, boolean inheritFromBefore) throws IOException {
        insertRowsColumns(Dimension.COLUMNS.getValue(), startIndex, endIndex, inheritFromBefore);
    }

    // =====================================
    // Builders
    // =====================================

    /**
     * Spreadsheet Builder
     */
    public static class Builder {
        private final Sheets service;
        private String spreadsheetId;
        private String range;

        /**
         * Constructor
         *
         * @param service an authorized Sheets API client service
         */
        public Builder(Sheets service) {
            this.service = service;
        }

        /**
         * Spreadsheet Id setter
         *
         * @param spreadsheetId the id of the spreadsheet
         * @return builder instance
         */
        public Builder withId(String spreadsheetId) {
            this.spreadsheetId = spreadsheetId;
            return this;
        }

        /**
         * Sheet range setter
         *
         * @param range the range
         * @return builder instance
         */
        public Builder inRange(String range) {
            this.range = range;
            return this;
        }

        /**
         * Build a SpreadSheet instance
         *
         * @return new instance of com.ydanchen.handysheets.apps.spreadsheet.SpreadSheet class
         */
        public SpreadSheet build() {
            return new SpreadSheet(this);
        }
    }

    // =====================================
    // Private methods
    // =====================================

    /**
     * @param dimension         What we want to insert. Rows or Columns
     * @param startIndex        start index
     * @param endIndex          end index
     * @param inheritFromBefore true to inherit
     * @throws IOException will be thrown if occurs
     */
    private void insertRowsColumns(String dimension, int startIndex, int endIndex, boolean inheritFromBefore) throws IOException {
        List<Request> requests = new ArrayList<>();
        DimensionRange range = new DimensionRange();
        range.setDimension(dimension);
        range.setStartIndex(startIndex);
        range.setEndIndex(endIndex);
        requests.add(new Request().setInsertDimension(new InsertDimensionRequest()
                        .setInheritFromBefore(inheritFromBefore)
                        .setRange(range)));
        BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest();
        requestBody.setRequests(requests);
        Sheets.Spreadsheets.BatchUpdate request = service.spreadsheets().batchUpdate(spreadsheetId, requestBody);
        BatchUpdateSpreadsheetResponse response = request.execute();
    }
}