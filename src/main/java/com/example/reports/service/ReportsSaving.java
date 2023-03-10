package com.example.reports.service;

import com.example.reports.config.AppConfig;
import com.example.reports.entity.Data;
import com.example.reports.entity.People;
import com.example.reports.entity.SecondFormPeople;
import com.example.reports.util.DepartmentUtil;
import com.example.reports.util.PathsConstants;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.math3.util.Pair;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import javax.annotation.PostConstruct;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Path;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

@Service
@RequiredArgsConstructor
@Slf4j
public class ReportsSaving {
    private static final String CURRENT_ENTRY = "current entry: {}";

    private static final String BEGIN_FILL = "begin fill: {}";
    private static final String DEPARTMENT_UTIL_BUT_NOT_RESUL = "DepartmentUtil has: {}, but result does not";

    private final List<ReportService> reports;
    private final AppConfig appConfig;
    private final DepartmentUtil departmentUtil;

    private List<Integer> postColumns;
    private List<Integer> bazaColumns;
    private List<Integer> dailyFirstColumns;
    private List<Integer> dailySecondColumns;
    private List<Integer> dailyThirdColumns;
    private List<Integer> fourteenColumns;
    private List<Integer> secondColumns;

    @PostConstruct
    void init() {
        postColumns = appConfig.getTo().getFirstForm().get("post");
        bazaColumns = appConfig.getTo().getFirstForm().get("baza");
        dailyFirstColumns = appConfig.getTo().getDailyForm().get("firstSheet");
        dailySecondColumns = appConfig.getTo().getDailyForm().get("secondSheet");
        dailyThirdColumns = appConfig.getTo().getDailyForm().get("thirdSheet");
        fourteenColumns = appConfig.getTo().getFourteenForm();
        secondColumns = appConfig.getTo().getSecondForm();
    }

    public void saveToFirstForm() throws Exception {
        log.info("saveToFirstForm() method invoked");
        Map<String, Data> post = getResult(FirstFormPostService.class);
        Map<String, Data> baza = getResult(FirstFormBazaService.class);
        try (FileInputStream file = new FileInputStream(PathsConstants.FIRST_TEMPLATE.toFile());
             FileOutputStream out = new FileOutputStream(Path.of(PathsConstants.templates, "0_30 forma.xlsx").toFile())){
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            fillBazaTable(baza, workbook);
            fillPostTable(post, workbook);

            workbook.write(out);
        }
    }

    public void saveToDailyForm() throws Exception {
        log.info("saveToDailyForm() method invoked");
        Map<String, Data> first = getResult(DailyFormFirstSheetService.class);
        Map<String, Data> secondAndThird = getResult(DailyFormSecondSheetService.class);
        try (FileInputStream file = new FileInputStream(PathsConstants.DAILY_TEMPLATE.toFile());
             FileOutputStream out = new FileOutputStream(Path.of(PathsConstants.templates, "0_dnevnoy.xlsx").toFile())) {
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            fillTableOne(first, workbook);
            fillTableTwo(secondAndThird, workbook);
            fillTableThree(secondAndThird, workbook);

            workbook.write(out);
        }
    }

    public void saveToFourteenForm() throws Exception {
        log.info("saveToFourteenForm() method invoked");
        Map<String, Data> fourteen = getResult(FourteenFormService.class);
        try (FileInputStream file = new FileInputStream(PathsConstants.FOURTEEN_TEMPLATE.toFile());
             FileOutputStream out = new FileOutputStream(Path.of(PathsConstants.templates, "0_14 forma.xlsx").toFile())){
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            String key;
            Integer value;
            Row row;
            log.debug(BEGIN_FILL, "?????????????? 2000_1");
            // Sheet "?????????????? 3100"
            XSSFSheet firstSheet = workbook.getSheet("??????????????2000_1");
            checkDepartmentUtilMapKeys(departmentUtil.getFourteenForma().keySet(), fourteen.keySet());
            for (Map.Entry<String, Integer> entry : departmentUtil.getFourteenForma().entrySet()) {
                log.debug(CURRENT_ENTRY, entry);
                value = entry.getValue();
                key = entry.getKey();
                row = firstSheet.getRow(value);
                if (fourteen.get(key) != null) {
                    setAdditionValue(row, fourteenColumns,
                            withoutDied(fourteen.get(key).getAdult()), // ???????????????????????????? ??????????
                            fourteen.get(key).getAdult().getEmergency(), // ???? ????????????????????
                            fourteen.get(key).getAdult().getAmbulance(), // ???????????????????? ????????????
                            fourteen.get(key).getAdult().getDays(), // ??????????????????
                            fourteen.get(key).getAdult().getDied(), // ????????????
                            withoutDied(fourteen.get(key).getOld()), // ???????????????????? ??????????
                            fourteen.get(key).getOld().getEmergency(), // ???? ????????????????????
                            fourteen.get(key).getOld().getAmbulance(), // ???????????????????? ????????????
                            fourteen.get(key).getOld().getDays(), // ??????????????????
                            fourteen.get(key).getOld().getDied(), // ????????????
                            withoutDied(fourteen.get(key).getChild()), // ???????? ??????????
                            fourteen.get(key).getChild().getEmergency(), // ???? ????????????????????
                            fourteen.get(key).getChild().getAmbulance(), // ???????????????????? ????????????
                            fourteen.get(key).getChild().getDays(), // ??????????????????
                            fourteen.get(key).getChild().getDied()); // ????????????
                    setAdditionValue(firstSheet.getRow(9), fourteenColumns, // ???????????????????? ???? ??????????
                            withoutDied(fourteen.get(key).getAdult()),
                            fourteen.get(key).getAdult().getEmergency(),
                            fourteen.get(key).getAdult().getAmbulance(),
                            fourteen.get(key).getAdult().getDays(),
                            fourteen.get(key).getAdult().getDied(),
                            withoutDied(fourteen.get(key).getOld()),
                            fourteen.get(key).getOld().getEmergency(),
                            fourteen.get(key).getOld().getAmbulance(),
                            fourteen.get(key).getOld().getDays(),
                            fourteen.get(key).getOld().getDied(),
                            withoutDied(fourteen.get(key).getChild()),
                            fourteen.get(key).getChild().getEmergency(),
                            fourteen.get(key).getChild().getAmbulance(),
                            fourteen.get(key).getChild().getDays(),
                            fourteen.get(key).getChild().getDied());
                } else {
                    log.debug(DEPARTMENT_UTIL_BUT_NOT_RESUL, key);
                }
            }

            groupByMKB(firstSheet);
            workbook.write(out);
        }
    }

    public void saveToSecondForm() throws Exception {
        log.info("saveToSecondForm() method invoked");
        for (ReportService service : reports) {
            if (service.getClass().equals(SecondFormService.class)) {
                Map<String, Pair<SecondFormPeople, SecondFormPeople>> result = ((SecondFormService) service).getSecondFormResult();
                try (FileInputStream file = new FileInputStream(PathsConstants.SECOND_TEMPLATE.toFile());
                     FileOutputStream out = new FileOutputStream(Path.of(PathsConstants.templates, "0_forma 2910.xlsx").toFile())) {
                    XSSFWorkbook workbook = new XSSFWorkbook(file);
                    XSSFSheet sheet = workbook.getSheet("??????????????2910");
                    Pair<SecondFormPeople, SecondFormPeople> pair = result.get("first");
                    setAdditionValue(sheet.getRow(6), secondColumns, convertPairToValues(pair));
                    pair = result.get("second");
                    setAdditionValue(sheet.getRow(7), secondColumns, convertPairToValues(pair));
                    pair = result.get("third");
                    setAdditionValue(sheet.getRow(8), secondColumns, convertPairToValues(pair));
                    pair = result.get("fourth");
                    setAdditionValue(sheet.getRow(9), secondColumns, convertPairToValues(pair));
                    pair = result.get("fifth");
                    setAdditionValue(sheet.getRow(10), secondColumns, convertPairToValues(pair));
                    pair = result.get("sixth");
                    setAdditionValue(sheet.getRow(11), secondColumns, convertPairToValues(pair));
                    pair = result.get("seventh");
                    setAdditionValue(sheet.getRow(12), secondColumns, convertPairToValues(pair));
                    workbook.write(out);
                }
                return;
            }
        }
    }

    private void fillBazaTable(Map<String, Data> baza, XSSFWorkbook workbook) {
        String key;
        Integer value;
        Row row;
        log.debug(BEGIN_FILL, "?????????????? 3100");
        // Sheet "?????????????? 3100"
        XSSFSheet firstSheet = workbook.getSheet("??????????????3100");
        checkDepartmentUtilMapKeys(departmentUtil.getFirstForma().keySet(), baza.keySet());
        for (Map.Entry<String, Integer> entry : departmentUtil.getFirstForma().entrySet()) {
            log.debug(CURRENT_ENTRY, entry);
            value = entry.getValue();
            key = entry.getKey();
            row = firstSheet.getRow(value);
            if (baza.get(key) != null) {
                setAdditionValue(row, bazaColumns,
                        withoutDied(baza.get(key).getAll()), // ???????????????? ?????????????????? ??????????
                        withoutDied(baza.get(key).getAdult()), // ????????????????????????????
                        withoutDied(baza.get(key).getChild()), // ????????
                        withoutDied(baza.get(key).getOld()), // ????????????
                        baza.get(key).getAll().getDied(), // ???????????? ??????????
                        baza.get(key).getAdult().getDied(), // ????????????????????????????
                        baza.get(key).getChild().getDied(), // ????????
                        baza.get(key).getOld().getDied(), // ????????????
                        baza.get(key).getAll().getDays(), // ?????????????????? ??????????
                        baza.get(key).getAdult().getDays(), // ????????????????????????????
                        baza.get(key).getChild().getDays(), // ????????
                        baza.get(key).getOld().getDays()); // ????????????
                setAdditionValue(firstSheet.getRow(4), bazaColumns, // ???????????????????? ???? ??????????
                        withoutDied(baza.get(key).getAll()),
                        withoutDied(baza.get(key).getAdult()),
                        withoutDied(baza.get(key).getChild()),
                        withoutDied(baza.get(key).getOld()),
                        baza.get(key).getAll().getDied(),
                        baza.get(key).getAdult().getDied(),
                        baza.get(key).getChild().getDied(),
                        baza.get(key).getOld().getDied(),
                        baza.get(key).getAll().getDays(),
                        baza.get(key).getAdult().getDays(),
                        baza.get(key).getChild().getDays(),
                        baza.get(key).getOld().getDays());
            } else {
                log.debug(DEPARTMENT_UTIL_BUT_NOT_RESUL, key);
            }
        }
    }

    private void fillPostTable(Map<String, Data> post, XSSFWorkbook workbook) {
        Integer value;
        String key;
        Row row;
        log.debug(BEGIN_FILL, "?????????????? 3100");
        // Sheet "?????????????? 3100"
        XSSFSheet firstSheet = workbook.getSheet("??????????????3100");
        checkDepartmentUtilMapKeys(departmentUtil.getFirstForma().keySet(), post.keySet());
        for (Map.Entry<String, Integer> entry : departmentUtil.getFirstForma().entrySet()) {
            log.debug(CURRENT_ENTRY, entry);
            value = entry.getValue();
            key = entry.getKey();
            row = firstSheet.getRow(value);
            if (post.get(key) != null) {
                setAdditionValue(row, postColumns,
                        post.get(key).getAll().getAll(), // ?????????????????? ?????????????????? - ??????????
                        post.get(key).getVillage(), // ???? ?????? ???????????????? ??????????????
                        post.get(key).getChild().getAll(), // 0-17 ?????? (????????????????????????)
                        post.get(key).getAdult().getAll(), // ????????????????????????????
                        post.get(key).getOld().getAll()); // ???????????? ?????????????????????????????? ????????????????
                setAdditionValue(firstSheet.getRow(4), postColumns, // ???????????????????? ???? ??????????
                        post.get(key).getAll().getAll(),
                        post.get(key).getVillage(),
                        post.get(key).getChild().getAll(),
                        post.get(key).getAdult().getAll(),
                        post.get(key).getOld().getAll());
            } else {
                log.debug(DEPARTMENT_UTIL_BUT_NOT_RESUL, key);
            }
        }
//        for (Map.Entry<String, Data> entry : post.entrySet()) {
//            log.debug(CURRENT_ENTRY, entry);
//            key = entry.getKey();
//            value = departmentUtil.getDailyFormaFirstSheet().get(key);
//            row = firstSheet.getRow(value);
//            setAdditionValue(row, postColumns,
//                    entry.getValue().getAll().getAll(), // ?????????????????? ?????????????????? - ??????????
//                    entry.getValue().getVillage(), // ???? ?????? ???????????????? ??????????????
//                    entry.getValue().getChild().getAll(), // 0-17 ?????? (????????????????????????)
//                    entry.getValue().getAdult().getAll(), // ????????????????????????????
//                    entry.getValue().getOld().getAll()); // ???????????? ?????????????????????????????? ????????????????
//            setAdditionValue(firstSheet.getRow(4), postColumns, // ???????????????????? ???? ??????????
//                    entry.getValue().getAll().getAll(),
//                    entry.getValue().getVillage(),
//                    entry.getValue().getChild().getAll(),
//                    entry.getValue().getAdult().getAll(),
//                    entry.getValue().getOld().getAll());
//        }
    }

    private void fillTableOne(Map<String, Data> first, XSSFWorkbook workbook) {
        String key;
        Integer value;
        Row row;
        log.debug(BEGIN_FILL, "?????????????? 2000");
        // Sheet "?????????????? 2000"
        XSSFSheet firstSheet = workbook.getSheet("??????????????2000");
        checkDepartmentUtilMapKeys(departmentUtil.getDailyFormaFirstSheet().keySet(), first.keySet());
        for (Map.Entry<String, Integer> entry : departmentUtil.getDailyFormaFirstSheet().entrySet()) {
            log.debug(CURRENT_ENTRY, entry);
            value = entry.getValue();
            key = entry.getKey();
            row = firstSheet.getRow(value);
            if (first.get(key) != null) {
                setCellValue(row, dailyFirstColumns,
                        (withoutDied(first.get(key).getAdult()) + withoutDied(first.get(key).getOld())), // ???????????????? ???????????????? ?? ??????????????????????
                        withoutDied(first.get(key).getOld()), // ???????????????? ??????????????????????
                        withoutDied(first.get(key).getChild()), // ???????????????? ??????????
                        (first.get(key).getAdult().getDays() + first.get(key).getOld().getDays()), // ?????????????????? ???????????????? ?? ??????????????????????
                        first.get(key).getOld().getDays(), // ?????????????????? ??????????????????????
                        first.get(key).getChild().getDays()); // ?????????????????? ??????????
                setAdditionValue(firstSheet.getRow(8), dailyFirstColumns, // ???????????????????? ???? ??????????
                        (withoutDied(first.get(key).getAdult()) + withoutDied(first.get(key).getOld())),
                        withoutDied(first.get(key).getOld()),
                        withoutDied(first.get(key).getChild()),
                        (first.get(key).getAdult().getDays() + first.get(key).getOld().getDays()),
                        first.get(key).getOld().getDays(),
                        first.get(key).getChild().getDays());
            } else {
                log.debug(DEPARTMENT_UTIL_BUT_NOT_RESUL, key);
            }
        }
    }

    private void fillTableTwo(Map<String, Data> secondAndThird, XSSFWorkbook workbook) {
        Row row;
        String key;
        Integer value;
        log.debug(BEGIN_FILL, "?????????????? 3000");
        // Sheet "?????????????? 3000"
        XSSFSheet secondSheet = workbook.getSheet("??????????????3000");
        checkDepartmentUtilMapKeys(departmentUtil.getDailyFormaSecondSheet().keySet(), secondAndThird.keySet());
        for (Map.Entry<String, Integer> entry : departmentUtil.getDailyFormaSecondSheet().entrySet()) {
            log.debug(CURRENT_ENTRY, entry);
            value = entry.getValue();
            key = entry.getKey();
            row = secondSheet.getRow(value);
            if (secondAndThird.get(key) != null) {
                setAdditionValue(row, dailySecondColumns,
                        (withoutDied(secondAndThird.get(key).getAdult()) + withoutDied(secondAndThird.get(key).getOld())), // ???????????????? ???????????????? ?? ??????????????????????
                        (secondAndThird.get(key).getAdult().getDays() + secondAndThird.get(key).getOld().getDays()), // ?????????????????? ???????????????? ?? ??????????????????????
                        (secondAndThird.get(key).getAdult().getDied() + secondAndThird.get(key).getOld().getDied())); // ????????????
                setAdditionValue(secondSheet.getRow(8), dailySecondColumns, // ???????????????????? ???? ??????????
                        (withoutDied(secondAndThird.get(key).getAdult()) + withoutDied(secondAndThird.get(key).getOld())),
                        (secondAndThird.get(key).getAdult().getDays() + secondAndThird.get(key).getOld().getDays()),
                        (secondAndThird.get(key).getAdult().getDied() + secondAndThird.get(key).getOld().getDied()));
            } else {
                log.debug(DEPARTMENT_UTIL_BUT_NOT_RESUL, key);
            }
        }
    }

    private void fillTableThree(Map<String, Data> secondAndThird, XSSFWorkbook workbook) {
        Integer value;
        String key;
        Row row;
        log.debug(BEGIN_FILL, "?????????????? 3500");
        // Sheet "?????????????? 3500"
        XSSFSheet thirdSheet = workbook.getSheet("??????????????3500");
        for (Map.Entry<String, Integer> entry : departmentUtil.getDailyFormaSecondSheet().entrySet()) {
            log.debug(CURRENT_ENTRY, entry);
            value = entry.getValue();
            key = entry.getKey();
            row = thirdSheet.getRow(value);
            if (secondAndThird.get(key) != null) {
                setAdditionValue(row, dailyThirdColumns,
                        withoutDied(secondAndThird.get(key).getChild()), // ???????????????? ??????????
                        secondAndThird.get(key).getChild().getDays(), // ?????????????????? ??????????
                        secondAndThird.get(key).getChild().getDied()); // ????????????
                setAdditionValue(thirdSheet.getRow(8), dailyThirdColumns, // ???????????????????? ???? ??????????
                        withoutDied(secondAndThird.get(key).getChild()),
                        secondAndThird.get(key).getChild().getDays(),
                        secondAndThird.get(key).getChild().getDied());
            } else {
                log.debug(DEPARTMENT_UTIL_BUT_NOT_RESUL, key);
            }
        }
    }

    private Integer[] convertPairToValues(Pair<SecondFormPeople, SecondFormPeople> pair) {
        Integer[] columnValues = new Integer[32];
        SecondFormPeople alive = pair.getFirst();
        SecondFormPeople dead = pair.getSecond();
        List<SecondFormPeople> both = List.of(alive, dead);
        for (int i = 0; i < 2; i++) {
            int add = 16 * i;
            columnValues[0 + add] = both.get(i).getBelow14();
            columnValues[1 + add] = both.get(i).getBetween15_19();
            columnValues[2 + add] = both.get(i).getBetween20_24();
            columnValues[3 + add] = both.get(i).getBetween25_29();
            columnValues[4 + add] = both.get(i).getBetween30_34();
            columnValues[5 + add] = both.get(i).getBetween35_39();
            columnValues[6 + add] = both.get(i).getBetween40_44();
            columnValues[7 + add] = both.get(i).getBetween45_49();
            columnValues[8 + add] = both.get(i).getBetween50_54();
            columnValues[9 + add] = both.get(i).getBetween55_59();
            columnValues[10 + add] = both.get(i).getBetween60_64();
            columnValues[11 + add] = both.get(i).getBetween65_69();
            columnValues[12 + add] = both.get(i).getBetween70_74();
            columnValues[13 + add] = both.get(i).getBetween75_79();
            columnValues[14 + add] = both.get(i).getBetween80_84();
            columnValues[15 + add] = both.get(i).getAbove85();
        }
        return columnValues;
    }

    private void checkDepartmentUtilMapKeys(Set<String> departmentUtilMapKeys, Set<String> resultMapKeys) {
        if (!departmentUtilMapKeys.containsAll(resultMapKeys)) {
            Set<String> temp = new HashSet<>(resultMapKeys);
            temp.removeAll(departmentUtilMapKeys);
            if (temp.size() == 1 && temp.contains("")) {
                return;
            }
            if (temp.size() == 1 && temp.contains(appConfig.getTempValue())) {
                return;
            }
            log.error("this keys are not present: {}", temp);
            throw new RuntimeException("Department Util Map does not have some keys");
        }
    }

    private Map<String, Data> getResult(Class<? extends ReportService> serviceClass) throws ClassNotFoundException {
        for (ReportService service : reports) {
            if (service.getClass().equals(serviceClass)) {
                return service.getResult();
            }
        }
        throw new ClassNotFoundException();
    }

    private Integer withoutDied(People people) {
        return people.getAll() - people.getDied();
    }

    /**
     * ???????????????????? ???????????????? ?? ????????????
     */
    private void setCellValue(Row row, List<Integer> columns, Integer... values) {
        for (int i = 0; i < columns.size(); i++) {
            row.getCell(columns.get(i)).setCellValue(values[i]);
        }
    }

    /**
     * ???????????????????? ?????????? ?????????????? ?? ???????????? ???????????????? ?? ???????????? (?????? ?????????? ?? "???? ??????")
     */
    private void setAdditionValue(Row row, List<Integer> columns, Integer... values) {
        for (int i = 0; i < columns.size(); i++) {
            row.getCell(columns.get(i)).setCellValue(row.getCell(columns.get(i)).getNumericCellValue() + values[i]);
        }
    }

    private void groupByMKB(Sheet sheet) {
        groupAB(sheet);
        groupCD4(sheet);
        groupD5(sheet);
        groupE(sheet);
        groupF(sheet);
        groupG(sheet);
        groupH(sheet);
        groupH6(sheet);
        groupI(sheet);
        groupJ(sheet);
        groupK(sheet);
        groupL(sheet);
        groupM(sheet);
        groupN(sheet);
        groupQ(sheet);
        groupR(sheet);
        groupST(sheet);
        subtractZ(sheet);
    }

    private void groupAB(Sheet sheet) {
        // A00-B99 2.0
        Integer[] rowsFrom = {11, 12, 13, 14, 15, 16, 17, 18, 19};
        Row rowTo = sheet.getRow(10);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupCD4(Sheet sheet) {
        // C84 3.1.1.7
        Integer[] rowsFrom = {30, 31};
        Row rowTo = sheet.getRow(29);
        groupRows(rowTo, rowsFrom, sheet);
        // C81-C96 3.1.1
        rowsFrom = new Integer[]{23, 24, 25, 26, 27, 28, 29, 32, 33, 34, 35, 36};
        rowTo = sheet.getRow(22);
        groupRows(rowTo, rowsFrom, sheet);
        // C00-C97 3.1
        rowsFrom = new Integer[]{22, 37};
        rowTo = sheet.getRow(21);
        groupRows(rowTo, rowsFrom, sheet);

        // D10-D36 3.3
        rowsFrom = new Integer[]{40, 41};
        rowTo = sheet.getRow(39);
        groupRows(rowTo, rowsFrom, sheet);

        // C00-D48 3.0
        rowsFrom = new Integer[]{21, 38, 39, 42};
        rowTo = sheet.getRow(20);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupD5(Sheet sheet) {
        // D50-D64 4.1
        Integer[] rowsFrom = {45, 46};
        Row rowTo = sheet.getRow(44);
        groupRows(rowTo, rowsFrom, sheet);
        // D65-D69 4.2
        rowsFrom = new Integer[]{48, 49};
        rowTo = sheet.getRow(47);
        groupRows(rowTo, rowsFrom, sheet);
        // D50-D89 4.0
        rowsFrom = new Integer[]{44, 47, 50, 51};
        rowTo = sheet.getRow(43);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupE(Sheet sheet) {
        // E10-E14 5.4
        Integer[] rowsFrom = {57, 58, 59, 60, 61};
        Row rowTo = sheet.getRow(56);
        groupRows(rowTo, rowsFrom, sheet);
        // E00-E89 5.0
        rowsFrom = new Integer[]{53, 54, 55, 56, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74};
        rowTo = sheet.getRow(52);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupF(Sheet sheet) {
        // F01-F99 6.0
        Integer[] rowsFrom = {76, 77};
        Row rowTo = sheet.getRow(75);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupG(Sheet sheet) {
        // G00-G09 7.1
        Integer[] rowsFrom = {80, 81, 82};
        Row rowTo = sheet.getRow(79);
        groupRows(rowTo, rowsFrom, sheet);
        // G20-G25 7.3
        rowsFrom = new Integer[]{85, 86, 87};
        rowTo = sheet.getRow(84);
        groupRows(rowTo, rowsFrom, sheet);
        // G30-G31 7.4
        rowsFrom = new Integer[]{89, 90};
        rowTo = sheet.getRow(88);
        groupRows(rowTo, rowsFrom, sheet);
        // G35-G37 7.5
        rowsFrom = new Integer[]{92, 93};
        rowTo = sheet.getRow(91);
        groupRows(rowTo, rowsFrom, sheet);
        // G40-G47 7.6
        rowsFrom = new Integer[]{95, 96, 97};
        rowTo = sheet.getRow(94);
        groupRows(rowTo, rowsFrom, sheet);
        // G50-G64 7.7
        rowsFrom = new Integer[]{99, 100};
        rowTo = sheet.getRow(98);
        groupRows(rowTo, rowsFrom, sheet);
        // G70-G73 7.8
        rowsFrom = new Integer[]{102, 103, 104};
        rowTo = sheet.getRow(101);
        groupRows(rowTo, rowsFrom, sheet);
        // G80-G83 7.9
        rowsFrom = new Integer[]{106, 107};
        rowTo = sheet.getRow(105);
        groupRows(rowTo, rowsFrom, sheet);
        // G00-G98 7.0
        rowsFrom = new Integer[]{79, 83, 84, 88, 91, 94, 98, 101, 105, 108, 109, 110};
        rowTo = sheet.getRow(78);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupH(Sheet sheet) {
        // H46-H48 8.8
        Integer[] rowsFrom = {120, 121};
        Row rowTo = sheet.getRow(119);
        groupRows(rowTo, rowsFrom, sheet);
        // H54 8.9
        rowsFrom = new Integer[]{123, 124};
        rowTo = sheet.getRow(122);
        groupRows(rowTo, rowsFrom, sheet);
        // H00-H59 8.0
        rowsFrom = new Integer[]{112, 113, 114, 115, 116, 117, 118, 119, 122, 125};
        rowTo = sheet.getRow(111);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupH6(Sheet sheet) {
        // H65-H74 9.1
        Integer[] rowsFrom = {128, 129, 130, 131, 132, 133};
        Row rowTo = sheet.getRow(127);
        groupRows(rowTo, rowsFrom, sheet);
        // H81-H83 9.2
        rowsFrom = new Integer[]{135, 136, 137};
        rowTo = sheet.getRow(134);
        groupRows(rowTo, rowsFrom, sheet);
        // H90 9.3
        rowsFrom = new Integer[]{139, 140, 141};
        rowTo = sheet.getRow(138);
        groupRows(rowTo, rowsFrom, sheet);
        // H60-H95 9.0
        rowsFrom = new Integer[]{127, 134, 138, 142};
        rowTo = sheet.getRow(126);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupI(Sheet sheet) {
        // I05-I09 10.2
        Integer[] rowsFrom = {146, 147};
        Row rowTo = sheet.getRow(145);
        groupRows(rowTo, rowsFrom, sheet);
        // I10-I13 10.3
        rowsFrom = new Integer[]{149, 150, 151, 152};
        rowTo = sheet.getRow(148);
        groupRows(rowTo, rowsFrom, sheet);
        // I20 10.4.1
        rowsFrom = new Integer[]{155, 156};
        rowTo = sheet.getRow(154);
        groupRows(rowTo, rowsFrom, sheet);
        // I25 10.4.5
        rowsFrom = new Integer[]{161, 162};
        rowTo = sheet.getRow(160);
        groupRows(rowTo, rowsFrom, sheet);
        // I20-I25 10.4
        rowsFrom = new Integer[]{154, 157, 158, 159, 160};
        rowTo = sheet.getRow(153);
        groupRows(rowTo, rowsFrom, sheet);
        // I30-I51 10.6
        rowsFrom = new Integer[]{165, 166, 167, 168, 169, 170, 171, 172, 173, 174};
        rowTo = sheet.getRow(164);
        groupRows(rowTo, rowsFrom, sheet);
        // I67 10.7.6
        rowsFrom = new Integer[]{182, 183};
        rowTo = sheet.getRow(181);
        groupRows(rowTo, rowsFrom, sheet);
        // I60-I69 10.7
        rowsFrom = new Integer[]{176, 177, 178, 179, 180, 181};
        rowTo = sheet.getRow(175);
        groupRows(rowTo, rowsFrom, sheet);
        // I80-I89 10.9
        rowsFrom = new Integer[]{186, 187, 188, 189};
        rowTo = sheet.getRow(185);
        groupRows(rowTo, rowsFrom, sheet);
        // I00-I99 10.0
        rowsFrom = new Integer[]{144, 145, 148, 153, 163, 164, 175, 184, 185, 190};
        rowTo = sheet.getRow(143);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupJ(Sheet sheet) {
        // J00-J06 11.1
        Integer[] rowsFrom = {193, 194, 195};
        Row rowTo = sheet.getRow(192);
        groupRows(rowTo, rowsFrom, sheet);
        // J00-J98 11.0
        rowsFrom = new Integer[]{192, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206};
        rowTo = sheet.getRow(191);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupK(Sheet sheet) {
        // K50-K52 12.4
        Integer[] rowsFrom = {212, 213, 214};
        Row rowTo = sheet.getRow(211);
        groupRows(rowTo, rowsFrom, sheet);
        // K55-K63 12.5
        rowsFrom = new Integer[]{216, 217, 218, 219, 220, 221};
        rowTo = sheet.getRow(215);
        groupRows(rowTo, rowsFrom, sheet);
        // K70-K76 12.8
        rowsFrom = new Integer[]{224, 225};
        rowTo = sheet.getRow(223);
        groupRows(rowTo, rowsFrom, sheet);
        // K85-K86 12.10
        rowsFrom = new Integer[]{228, 229};
        rowTo = sheet.getRow(227);
        groupRows(rowTo, rowsFrom, sheet);
        // K00-K92 12.0
        rowsFrom = new Integer[]{208, 209, 210, 211, 215, 222, 223, 226, 227, 230};
        rowTo = sheet.getRow(207);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupL(Sheet sheet) {
        // L40 13.4
        Integer[] rowsFrom = {236, 237};
        Row rowTo = sheet.getRow(235);
        groupRows(rowTo, rowsFrom, sheet);
        // L00-L98 13.0
        rowsFrom = new Integer[]{232, 233, 234, 235, 238, 239, 240};
        rowTo = sheet.getRow(231);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupM(Sheet sheet) {
        // M00-M25 14.1
        Integer[] rowsFrom = {243, 244, 245, 246, 247};
        Row rowTo = sheet.getRow(242);
        groupRows(rowTo, rowsFrom, sheet);
        // M30-M35 14.2
        rowsFrom = new Integer[]{249, 250};
        rowTo = sheet.getRow(248);
        groupRows(rowTo, rowsFrom, sheet);
        // M45-M49 14.4
        rowsFrom = new Integer[]{253, 254};
        rowTo = sheet.getRow(252);
        groupRows(rowTo, rowsFrom, sheet);
        // M80-M94 14.7
        rowsFrom = new Integer[]{258, 259};
        rowTo = sheet.getRow(257);
        groupRows(rowTo, rowsFrom, sheet);
        // M00-M99 14.0
        rowsFrom = new Integer[]{242, 248, 251, 252, 255, 256, 257, 260};
        rowTo = sheet.getRow(241);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupN(Sheet sheet) {
        // N70-N76 15.7
        Integer[] rowsFrom = {269, 270};
        Row rowTo = sheet.getRow(268);
        groupRows(rowTo, rowsFrom, sheet);
        // N00-N99 15.0
        rowsFrom = new Integer[]{262, 263, 264, 265, 266, 267, 268, 271, 272, 273, 274, 275};
        rowTo = sheet.getRow(261);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupQ(Sheet sheet) {
        // Q38-Q45 18.4
        Integer[] rowsFrom = {283, 284};
        Row rowTo = sheet.getRow(282);
        groupRows(rowTo, rowsFrom, sheet);
        // Q00-Q99 18.0
        rowsFrom = new Integer[]{279, 280, 281, 282, 285, 286, 287, 288, 289, 290};
        rowTo = sheet.getRow(278);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupR(Sheet sheet) {
        // R16 - R17
        Integer[] rowsFrom = {301};
        Row rowTo = sheet.getRow(291);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupST(Sheet sheet) {
        // S*2-T14 20.1
        Integer[] rowsFrom = {343, 344};
        Row rowTo = sheet.getRow(342);
        groupRows(rowTo, rowsFrom, sheet);
        // T36-T50 20.5
        rowsFrom = new Integer[]{349, 350};
        rowTo = sheet.getRow(348);
        groupRows(rowTo, rowsFrom, sheet);
        // T51-T65 20.6
        rowsFrom = new Integer[]{352, 353};
        rowTo = sheet.getRow(351);
        groupRows(rowTo, rowsFrom, sheet);
        // Soo-T98 20.0
        rowsFrom = new Integer[]{342, 345, 346, 347, 348, 351, 354};
        rowTo = sheet.getRow(341);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupRows(Row rowTo, Integer[] rowsFrom, Sheet sheet) {
        for (Integer row : rowsFrom) {
            Row from = sheet.getRow(row);
            for (Integer column : fourteenColumns) {
                double valueFrom = from.getCell(column).getNumericCellValue();
                double valueTo = rowTo.getCell(column).getNumericCellValue();
                rowTo.getCell(column).setCellValue(valueTo + valueFrom);
            }
        }
    }

    private void subtractZ(Sheet sheet) {
        Row from = sheet.getRow(9); // ??????????
        Row z = sheet.getRow(355); // Z00-Z99 21.0
        for (Integer column : fourteenColumns) {
            double valueFrom = from.getCell(column).getNumericCellValue();
            double valueZ = z.getCell(column).getNumericCellValue();
            from.getCell(column).setCellValue(valueFrom - valueZ);
        }
    }
}