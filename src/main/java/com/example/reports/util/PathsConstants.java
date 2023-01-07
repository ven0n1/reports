package com.example.reports.util;

import java.net.URISyntaxException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Objects;

public class PathsConstants {
    public static String src;
    public static final String BAZA = "baza.xlsx";
    public static final String POST = "post.xlsx";
    public static final String TEMP = "temp.csv";
    public static String templates;

    static {
        try {
            src = Paths.get(Objects.requireNonNull(
                    PathsConstants.class.getClassLoader().getResource("origin")).toURI()).toString();
            templates = Paths.get(Objects.requireNonNull(
                    PathsConstants.class.getClassLoader().getResource("templates")).toURI()).toString();
        } catch (URISyntaxException e) {
            e.printStackTrace();
        }
    }

    public static final Path FROM_BAZA = Path.of(src, BAZA);
    public static final Path FROM_POST = Path.of(src, POST);
    public static final Path FROM_TEMP = Path.of(src, TEMP);
    public static final Path FOR_DAILY = Path.of(src, "for_daily");
    public static final Path FOR_DAILY_FIRST = FOR_DAILY.resolve("dailyFirst.xlsx");
    public static final Path FOR_DAILY_SECOND = FOR_DAILY.resolve("dailySecond.xlsx");
    public static final Path FOR_FIRST = Path.of(src, "for_first");
    public static final Path FOR_FIRST_BAZA = FOR_FIRST.resolve(BAZA);
    public static final Path FOR_FIRST_POST = FOR_FIRST.resolve(POST);
    public static final Path FOR_FOURTEEN = Path.of(src, "for_fourteen");
    public static final Path FOR_FOURTEEN_BAZA = FOR_FOURTEEN.resolve("fourteen.xlsx");
    public static final Path DAILY_TEMPLATE = Path.of(templates, "daily_form.xlsx");
    public static final Path FIRST_TEMPLATE = Path.of(templates, "first_form.xlsx");
    public static final Path FOURTEEN_TEMPLATE = Path.of(templates, "fourteen_form.xlsx");

    private PathsConstants() {}
}
