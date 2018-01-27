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
    private static final char CHAR_A = 'A';
    private static final char CHAR_Z = 'Z';
    private static final char CHAR_0 = '0';
    private static final char CHAR_9 = '9';
    private static final int ALPHABET_LENGTH = 26;

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
     * Converts a List to a two dimensional array of Objects with a size of [n][1]
     * Useful when you need to write one column to a spreadsheet
     *
     * @param values the Column data as List
     * @return two dimensional array of Objects
     */
    public static Object[][] columnToTwoDimArray(List<Object> values) {
        Object[][] data = new Object[values.size()][1];
        for (int i = 0; i < values.size(); i++)
            data[i][0] = values.get(i);
        return data;
    }

    /**
     * Converts a List to a two dimensional array of Objects with a size of [1][n]
     * Useful when you need to write one row to a spreadsheet
     *
     * @param values the Row data as List
     * @return two dimensional array of Objects
     */
    public static Object[][] rowToTwoDimArray(List<Object> values) {
        Object[][] data = new Object[1][values.size()];
        for (int i = 0; i < values.size(); i++)
            data[0][i] = values.get(i);
        return data;
    }

    /**
     * Converts numeric indexes to a literal range, e.g. (1, 1, 2, 2) --> A1:B2
     *
     * @param startColumn start column index
     * @param startRow    start row index
     * @param endColumn   end column index
     * @param endRow      end row index
     * @return the formatted range
     */
    public static String numericRangeToLiteral(int startColumn, int startRow, int endColumn, int endRow) {
        return columnIndexToLetter(startColumn) + startRow + ":" + columnIndexToLetter(endColumn) + endRow;
    }

    /**
     * Converts numeric range instance to a literal
     *
     * @param range {@link NumericRange}
     * @return the formatted range, e.g. 'A1:C3'
     */
    public static String numericRangeToLiteral(NumericRange range) {
        return numericRangeToLiteral(range.getStartColumn(),
                range.getStartRow(),
                range.getEndColumn(),
                range.getEndRow());
    }

    /**
     * Converts literal range to a numeric
     *
     * @param range formatted range like 'A1:C3"
     * @return the instance of {@link NumericRange}
     */
    public static NumericRange literalRangeToNumerical(String range) {
        String[] tokens = range.toUpperCase().split(":");
        int startColumn = letterToColumnIndex(getLetterPart(tokens[0]));
        int startRow = Integer.parseInt(getNumericPart(tokens[0]));
        int endColumn = letterToColumnIndex(getLetterPart(tokens[1]));
        int endRow = Integer.parseInt(getNumericPart(tokens[1]));
        return new NumericRange(startColumn, startRow, endColumn, endRow);
    }

    // =================
    //  Private Methods
    // =================

    /**
     * Converts column index to a letter representation, e.g. (1 -> A) or (30 -> AD)
     *
     * @param index index of the column
     * @return the letter representation
     */
    private static String columnIndexToLetter(int index) {
        int temp;
        StringBuilder letter = new StringBuilder();
        while (index > 0) {
            temp = (index - 1) % ALPHABET_LENGTH;
            letter.insert(0, (char) (temp + CHAR_A));
            index = (index - temp - 1) / ALPHABET_LENGTH;
        }
        return letter.toString();
    }

    /**
     * Converts column literal to an index, e.g. (A -> 1) or (30 -> AD)
     *
     * @param literal column name
     * @return the numerical index
     */
    private static int letterToColumnIndex(String literal) {
        int number = 0;
        for (int i = 0; i < literal.length(); i++) {
            number = number * ALPHABET_LENGTH + (literal.charAt(i) - (CHAR_A - 1));
        }
        return number;
    }

    /**
     * Gets a numeric part from a range token
     * e.g. (A100 -> 100) or (B4 -> 4)
     *
     * @param token the part of a range,
     *              where [A1:B2] is a range and [A1] and [B2] are the range tokens
     * @return the numeric part of the token
     */
    private static String getNumericPart(String token) {
        return token.chars()
                .filter(s -> s >= CHAR_0 && s <= CHAR_9)
                .mapToObj(s -> (char) s)
                .collect(StringBuilder::new, StringBuilder::append, StringBuilder::append)
                .toString();
    }

    /**
     * Gets a letter part from a range token
     * e.g. (A100 -> A) or (AZ14 -> AZ)
     *
     * @param token the part of a range,
     *              where [A1:B2] is a range and [A1] and [B2] are the range tokens
     * @return the literal part of the token
     */
    private static String getLetterPart(String token) {
        return token.chars()
                .filter(s -> s >= CHAR_A && s <= CHAR_Z)
                .mapToObj(s -> (char) s)
                .collect(StringBuilder::new, StringBuilder::append, StringBuilder::append)
                .toString();
    }
}