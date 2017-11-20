package com.ydanchen.handysheet;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import com.ydanchen.handysheet.enums.Dimension;
import com.ydanchen.handysheet.enums.InputOptionValue;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

/**
 * This class provides access to the most common Google SpreadSheet operations
 *
 * @author Yevhen Danchenko
 */
public class SpreadSheet {
    private Sheets service;
    private String spreadsheetId;
    private String range;
    private InputOptionValue inputOptionValue;

    /**
     * Constructor
     *
     * @param builder {@link Builder}
     */
    private SpreadSheet(Builder builder) {
        this.service = builder.service;
        this.spreadsheetId = builder.spreadsheetId;
        this.range = builder.range;
        this.inputOptionValue = builder.inputOptionValue;
    }

    /**
     * Range setter
     *
     * @param range the new Range
     */
    public SpreadSheet setRange(final String range) {
        this.range = range;
        return this;
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
        //TODO: TBD
        return null;
    }

    /**
     * Update values on the spreadsheet from the List of Lists
     * Range should be specified before
     *
     * @param values the values to set
     * @return {@link UpdateValuesResponse}
     * @throws IOException might be thrown
     */
    public UpdateValuesResponse updateValues(final List<List<Object>> values) throws IOException {
        return updateValuesApiCall(values);
    }

    /**
     * Update values on the spreadsheet from the two dimensional array
     *
     * @param values the values to set
     * @return {@link UpdateValuesResponse}
     * @throws IOException might be thrown
     */
    public UpdateValuesResponse updateValues(Object[][] values) throws IOException {
        List<List<Object>> list = Arrays.stream(values)
                .map(Arrays::asList)
                .collect(Collectors.toList());
        return updateValuesApiCall(list);
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
        return insertRowsColumnsApiCall(Dimension.ROWS.getValue(), startIndex, endIndex, inheritFromBefore);
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
        return insertRowsColumnsApiCall(Dimension.COLUMNS.getValue(), startIndex, endIndex, inheritFromBefore);
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
        private InputOptionValue inputOptionValue;

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
         * Input Option Value setter
         *
         * @param inputOptionValue the Input Option
         * @return builder instance
         */
        public Builder withInputOptionValue(InputOptionValue inputOptionValue) {
            this.inputOptionValue = inputOptionValue;
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
     * Inserts Rows or Columns to the spreadsheet
     *
     * @param dimension         What we want to insert. Rows or Columns
     * @param startIndex        start index
     * @param endIndex          end index
     * @param inheritFromBefore true to inherit
     * @return {@link BatchUpdateSpreadsheetResponse}
     * @throws IOException will be thrown if occurs
     */
    private BatchUpdateSpreadsheetResponse insertRowsColumnsApiCall(String dimension, int startIndex, int endIndex, boolean inheritFromBefore) throws IOException {
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
        return request.execute();
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
        return service.spreadsheets().values().update(spreadsheetId, range, body)
                .setValueInputOption(inputOptionValue.getValue())
                .execute();
    }
}