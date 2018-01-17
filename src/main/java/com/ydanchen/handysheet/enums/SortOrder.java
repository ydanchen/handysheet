package com.ydanchen.handysheet.enums;

/**
 * A sort order enum
 *
 * @author Yevhen Danchenko
 */
public enum SortOrder {
    /**
     * Sort ascending
     */
    ASCENDING("ASCENDING"),
    /**
     * Sort descending
     */
    DESCENDING("DESCENDING");

    private final String value;

    /**
     * Constructor
     *
     * @param value the value of the enum
     */
    SortOrder(final String value) {
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