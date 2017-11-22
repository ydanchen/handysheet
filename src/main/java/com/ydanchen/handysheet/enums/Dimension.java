package com.ydanchen.handysheet.enums;

/**
 * Spreadsheet Dimensions enum
 *
 * @author Yevhen Danchenko
 * @link https://developers.google.com/sheets/api/reference/rest/v4/Dimension
 */
public enum Dimension {
    /**
     * Represents Rows of the spreadsheet
     */
    ROWS("ROWS"),
    /**
     * Represents Columns of the spreadsheet
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