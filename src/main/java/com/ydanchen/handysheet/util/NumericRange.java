package com.ydanchen.handysheet.util;

import java.util.Objects;

/**
 * This class represents the Range on the spreadsheet
 */
public class NumericRange {
    private int startColumn;
    private int startRow;
    private int endColumn;
    private int endRow;

    /**
     * Constructor
     *
     * @param startColumn the start column of the range
     * @param startRow    the start row of the range
     * @param endColumn   the end column of the range
     * @param endRow      the end row oth range
     */
    public NumericRange(int startColumn, int startRow, int endColumn, int endRow) {
        this.startColumn = startColumn;
        this.startRow = startRow;
        this.endColumn = endColumn;
        this.endRow = endRow;
    }

    /**
     * Start Column getter
     *
     * @return the start Column
     */
    public int getStartColumn() {
        return startColumn;
    }

    /**
     * Start Row getter
     *
     * @return the start Row
     */
    public int getStartRow() {
        return startRow;
    }

    /**
     * End Column getter
     *
     * @return the end Column
     */
    public int getEndColumn() {
        return endColumn;
    }

    /**
     * End Row getter
     *
     * @return the end Row
     */
    public int getEndRow() {
        return endRow;
    }

    // =================
    // Override Methods
    // =================

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        NumericRange that = (NumericRange) o;
        return startColumn == that.startColumn &&
                startRow == that.startRow &&
                endColumn == that.endColumn &&
                endRow == that.endRow;
    }

    @Override
    public int hashCode() {
        return Objects.hash(startColumn, startRow, endColumn, endRow);
    }

    @Override
    public String toString() {
        return startColumn + ", " + startRow + ", " + endColumn + ", " + endRow;
    }
}
