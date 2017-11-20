package com.ydanchen.handysheet.enums;

/**
 * Spreadsheet Dimensions enum
 *
 * @link https://developers.google.com/sheets/api/reference/rest/v4/Dimension
 */
public enum Dimension {
    /**
     * Rows
     */
    ROWS("ROWS"),

    /**
     * Columns
     */
    COLUMNS("COLUMNS");

    private final String value;

    /**
     * Constructor
     *
     * @param value the value of the enum
     */
    Dimension(final String value) {
        this.value = value;
    }

    /**
     * Gets string representation of the enum
     *
     * @return string representation
     */
    public String getValue() {
        return this.value;
    }
}