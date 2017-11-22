package com.ydanchen.handysheet.util;

import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Some utility methods: type conversion etc.
 *
 * @author Yevhen Danchenko
 */
public class Utils {
    /**
     * Converts two dimensional array of Objects to List of Lists of Objects
     *
     * @param values two dimensional array of Objects
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
     * @param values List of Lists of Objects
     * @return two dimensional array of Objects
     */
    public static Object[][] listOfListsToTwoDimArray(List<List<Object>> values) {
        return values.stream()
                .map(l -> l.stream().toArray(Object[]::new))
                .toArray(Object[][]::new);
    }
}
