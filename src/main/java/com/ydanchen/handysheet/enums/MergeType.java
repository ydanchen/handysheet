package com.ydanchen.handysheet.enums;

/**
 * The type of merge to create.
 *
 * @author Yevhen Danchenko
 * @link https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#mergetype
 */
public enum MergeType {
    /**
     * Create a single merge from the range
     */
    MERGE_ALL("MERGE_ALL"),
    /**
     * Create a merge for each column in the range
     */
    MERGE_COLUMNS("MERGE_COLUMNS"),
    /**
     * Create a merge for each row in the range
     */
    MERGE_ROWS("MERGE_ROWS");

    private final String value;

    /**
     * Constructor
     *
     * @param value the value of the enum
     */
    MergeType(final String value) {
        this.value = value;
    }

    /**
     * Gets string representation of the enum
     *
     * @return string representation
     */
    public String getValue() {
        return value;
    }
}