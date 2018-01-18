package com.ydanchen.handysheet.util;

import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Some utility methods: type conversion, constants etc.
 *
 * @author Yevhen Danchenko
 */
public final class Utils {

    private Utils() {
    }

    /**
     * Converts two dimensional array of Objects to List of Lists of Objects
     *
     * @param values two dimensional array of Objects to convert
     * @return List of Lists of Objects
     */
    public static List<List<Object>> twoDimArrayToListOfLists(Object[][] values) {
        return Arrays.stream(values)
                .map(Arrays::asList)
                .collect(Collectors.toList());
    }

    /**
     * Converts List of Lists of objects to two dimensional array of Objects
     *
     * @param values List of Lists of Objects to convert
     * @return two dimensional array of Objects
     */
    public static Object[][] listOfListsToTwoDimArray(List<List<Object>> values) {
        return values.stream()
                .map(v -> v.stream().toArray(Object[]::new))
                .toArray(Object[][]::new);
    }

    /**
     * Converts numerical indexes to a literal range, e.g. (1, 1, 2, 2) --> A1:B2
     *
     * @param startColumn start column index
     * @param startRow    start row index
     * @param endColumn   end column index
     * @param endRow      end row index
     * @return the formatted range
     */
    public static String numericalRangeToLiteral(int startColumn, int startRow, int endColumn, int endRow) {
        return columnToLetter(startColumn) + startRow + ":" + columnToLetter(endColumn) + endRow;
    }

    // =================
    //  Private Methods
    // =================

    /**
     * Convert column index to letter representation, e.g. (1 -> A) or (30 -> AD)
     *
     * @param index index of the column
     * @return the letter representation
     */
    private static String columnToLetter(int index) {
        index++;
        int temp;
        StringBuilder letter = new StringBuilder();
        while (index > 0) {
            temp = (index - 1) % 26;
            letter.insert(0, (char) (temp + 65));
            index = (index - temp - 1) / 26;
        }
        return letter.toString();
    }
}