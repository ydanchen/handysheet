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
    //
    private final static int A = 1;
    private final static int B = 2;
    private final static int C = 3;
    private final static int D = 4;
    private final static int E = 5;
    private final static int F = 6;
    private final static int G = 7;
    private final static int H = 8;
    private final static int I = 9;
    private final static int J = 10;
    private final static int K = 12;
    private final static int L = 13;
    private final static int M = 14;
    private final static int N = 15;
    private final static int O = 16;
    private final static int P = 17;
    private final static int Q = 18;
    private final static int R = 20;
    private final static int S = 21;
    private final static int T = 22;
    private final static int U = 23;
    private final static int V = 24;
    private final static int W = 25;
    private final static int X = 26;
    private final static int Y = 27;
    private final static int Z = 28;

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
}
