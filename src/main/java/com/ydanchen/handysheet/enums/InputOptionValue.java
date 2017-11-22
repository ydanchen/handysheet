package com.ydanchen.handysheet.enums;

/**
 * Spreadsheet Input Option Value
 *
 * @author Yevhen Danchenko
 * @link https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption
 */
public enum InputOptionValue {
    /**
     * The values the user has entered will not be parsed and will be stored as-is.
     */
    RAW("RAW"),

    /**
     * The values will be parsed as if the user typed them into the UI.
     * Numbers will stay as numbers, but strings may be converted to numbers, dates, etc.
     * following the same rules that are applied when entering text into a cell via the Google Sheets UI.
     */
    USER_ENTERED("USER_ENTERED");

    private final String value;

    /**
     * Constructor
     *
     * @param value the value of the enum
     */
    InputOptionValue(final String value) {
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