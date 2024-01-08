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
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

@Service
@RequiredArgsConstructor
@Slf4j
public class ReportsSaving {
    private static final String CURRENT_ENTRY = "current entry: {}";

    private static final String BEGIN_FILL = "begin fill: {}";

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
            log.debug(BEGIN_FILL, "Таблица 2000_1");
            // Sheet "Таблица 3100"
            XSSFSheet firstSheet = workbook.getSheet("Таблица2000_1");
            checkDepartmentUtilMapKeys(departmentUtil.getFourteenForma().keySet(), fourteen.keySet());
            for (Map.Entry<String, Integer> entry : departmentUtil.getFourteenForma().entrySet()) {
                log.debug(CURRENT_ENTRY, entry);
                value = entry.getValue();
                key = entry.getKey();
                row = firstSheet.getRow(value);
                if (fourteen.get(key) != null) {
                    setAdditionValue(row, fourteenColumns,
                            withoutDied(fourteen.get(key).getAdult()), // ТРУДОСПОСОБНЫЕ всего
                            fourteen.get(key).getAdult().getEmergency(), // по экстренным
                            fourteen.get(key).getAdult().getAmbulance(), // доставлены скорой
                            fourteen.get(key).getAdult().getDays(), // проведено
                            fourteen.get(key).getAdult().getDied(), // умерло
                            withoutDied(fourteen.get(key).getOld()), // ПЕНСИОНЕРЫ всего
                            fourteen.get(key).getOld().getEmergency(), // по экстренным
                            fourteen.get(key).getOld().getAmbulance(), // доставлены скорой
                            fourteen.get(key).getOld().getDays(), // проведено
                            fourteen.get(key).getOld().getDied(), // умерло
                            withoutDied(fourteen.get(key).getChild()), // ДЕТИ всего
                            fourteen.get(key).getChild().getEmergency(), // по экстренным
                            fourteen.get(key).getChild().getAmbulance(), // доставлены скорой
                            fourteen.get(key).getChild().getDays(), // проведено
                            fourteen.get(key).getChild().getDied()); // умерло
                    setAdditionValue(firstSheet.getRow(9), fourteenColumns, // Сохранение во ВСЕГО
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
                    XSSFSheet sheet = workbook.getSheet("Таблица2910");
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
        log.debug(BEGIN_FILL, "Таблица 3100");
        // Sheet "Таблица 3100"
        XSSFSheet firstSheet = workbook.getSheet("Таблица3100");
        checkDepartmentUtilMapKeys(departmentUtil.getFirstForma().keySet(), baza.keySet());
        for (Map.Entry<String, Integer> entry : departmentUtil.getFirstForma().entrySet()) {
            log.debug(CURRENT_ENTRY, entry);
            value = entry.getValue();
            key = entry.getKey();
            row = firstSheet.getRow(value);
            if (baza.get(key) != null) {
                setAdditionValue(row, bazaColumns,
                        withoutDied(baza.get(key).getAll()), // ВЫПИСАНО пациентов всего
                        withoutDied(baza.get(key).getAdult()), // трудоспособные
                        withoutDied(baza.get(key).getChild()), // дети
                        withoutDied(baza.get(key).getOld()), // старше
                        baza.get(key).getAll().getDied(), // УМЕРЛО всего
                        baza.get(key).getAdult().getDied(), // трудоспособные
                        baza.get(key).getChild().getDied(), // дети
                        baza.get(key).getOld().getDied(), // старше
                        baza.get(key).getAll().getDays(), // ПРОВЕДЕНО всего
                        baza.get(key).getAdult().getDays(), // трудоспособные
                        baza.get(key).getChild().getDays(), // дети
                        baza.get(key).getOld().getDays()); // старше
                setAdditionValue(firstSheet.getRow(4), bazaColumns, // Сохранение во ВСЕГО
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
            }
        }
    }

    private void fillPostTable(Map<String, Data> post, XSSFWorkbook workbook) {
        Integer value;
        String key;
        Row row;
        log.debug(BEGIN_FILL, "Таблица 3100");
        // Sheet "Таблица 3100"
        XSSFSheet firstSheet = workbook.getSheet("Таблица3100");
        checkDepartmentUtilMapKeys(departmentUtil.getFirstForma().keySet(), post.keySet());
        for (Map.Entry<String, Integer> entry : departmentUtil.getFirstForma().entrySet()) {
            log.debug(CURRENT_ENTRY, entry);
            value = entry.getValue();
            key = entry.getKey();
            row = firstSheet.getRow(value);
            if (post.get(key) != null) {
                setAdditionValue(row, postColumns,
                        post.get(key).getAll().getAll(), // Поступило пациентов - всего
                        post.get(key).getVillage(), // из них сельских жителей
                        post.get(key).getChild().getAll(), // 0-17 лет (включительно)
                        post.get(key).getAdult().getAll(), // трудоспособные
                        post.get(key).getOld().getAll()); // старше трудоспособного возраста
                setAdditionValue(firstSheet.getRow(4), postColumns, // Сохранение во ВСЕГО
                        post.get(key).getAll().getAll(),
                        post.get(key).getVillage(),
                        post.get(key).getChild().getAll(),
                        post.get(key).getAdult().getAll(),
                        post.get(key).getOld().getAll());
            }
        }
//        for (Map.Entry<String, Data> entry : post.entrySet()) {
//            log.debug(CURRENT_ENTRY, entry);
//            key = entry.getKey();
//            value = departmentUtil.getDailyFormaFirstSheet().get(key);
//            row = firstSheet.getRow(value);
//            setAdditionValue(row, postColumns,
//                    entry.getValue().getAll().getAll(), // Поступило пациентов - всего
//                    entry.getValue().getVillage(), // из них сельских жителей
//                    entry.getValue().getChild().getAll(), // 0-17 лет (включительно)
//                    entry.getValue().getAdult().getAll(), // трудоспособные
//                    entry.getValue().getOld().getAll()); // старше трудоспособного возраста
//            setAdditionValue(firstSheet.getRow(4), postColumns, // Сохранение во ВСЕГО
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
        log.debug(BEGIN_FILL, "Таблица 2000");
        // Sheet "Таблица 2000"
        XSSFSheet firstSheet = workbook.getSheet("Таблица2000");
        checkDepartmentUtilMapKeys(departmentUtil.getDailyFormaFirstSheet().keySet(), first.keySet());
        for (Map.Entry<String, Integer> entry : departmentUtil.getDailyFormaFirstSheet().entrySet()) {
            log.debug(CURRENT_ENTRY, entry);
            value = entry.getValue();
            key = entry.getKey();
            row = firstSheet.getRow(value);
            if (first.get(key) != null) {
                setCellValue(row, dailyFirstColumns,
                        (withoutDied(first.get(key).getAdult()) + withoutDied(first.get(key).getOld())), // Выписано взрослых и пенсионеров
                        withoutDied(first.get(key).getOld()), // Выписано пенсионеров
                        withoutDied(first.get(key).getChild()), // Выписано детей
                        (first.get(key).getAdult().getDays() + first.get(key).getOld().getDays()), // Проведено взрослых и пенсионеров
                        first.get(key).getOld().getDays(), // Проведено пенсионеров
                        first.get(key).getChild().getDays()); // Проведено детей
                setAdditionValue(firstSheet.getRow(8), dailyFirstColumns, // Сохранение во ВСЕГО
                        (withoutDied(first.get(key).getAdult()) + withoutDied(first.get(key).getOld())),
                        withoutDied(first.get(key).getOld()),
                        withoutDied(first.get(key).getChild()),
                        (first.get(key).getAdult().getDays() + first.get(key).getOld().getDays()),
                        first.get(key).getOld().getDays(),
                        first.get(key).getChild().getDays());
            }
        }
    }

    private void fillTableTwo(Map<String, Data> secondAndThird, XSSFWorkbook workbook) {
        Row row;
        String key;
        Integer value;
        log.debug(BEGIN_FILL, "Таблица 3000");
        // Sheet "Таблица 3000"
        XSSFSheet secondSheet = workbook.getSheet("Таблица3000");
        checkDepartmentUtilMapKeys(departmentUtil.getDailyFormaSecondSheet().keySet(), secondAndThird.keySet());
        for (Map.Entry<String, Integer> entry : departmentUtil.getDailyFormaSecondSheet().entrySet()) {
            log.debug(CURRENT_ENTRY, entry);
            value = entry.getValue();
            key = entry.getKey();
            row = secondSheet.getRow(value);
            if (secondAndThird.get(key) != null) {
                setAdditionValue(row, dailySecondColumns,
                        (withoutDied(secondAndThird.get(key).getAdult()) + withoutDied(secondAndThird.get(key).getOld())), // Выписано взрослых и пенсионеров
                        (secondAndThird.get(key).getAdult().getDays() + secondAndThird.get(key).getOld().getDays()), // Проведено взрослых и пенсионеров
                        (secondAndThird.get(key).getAdult().getDied() + secondAndThird.get(key).getOld().getDied())); // Умерло
                setAdditionValue(secondSheet.getRow(8), dailySecondColumns, // Сохранение во ВСЕГО
                        (withoutDied(secondAndThird.get(key).getAdult()) + withoutDied(secondAndThird.get(key).getOld())),
                        (secondAndThird.get(key).getAdult().getDays() + secondAndThird.get(key).getOld().getDays()),
                        (secondAndThird.get(key).getAdult().getDied() + secondAndThird.get(key).getOld().getDied()));
            }
        }
    }

    private void fillTableThree(Map<String, Data> secondAndThird, XSSFWorkbook workbook) {
        Integer value;
        String key;
        Row row;
        log.debug(BEGIN_FILL, "Таблица 3500");
        // Sheet "Таблица 3500"
        XSSFSheet thirdSheet = workbook.getSheet("Таблица3500");
        for (Map.Entry<String, Integer> entry : departmentUtil.getDailyFormaSecondSheet().entrySet()) {
            log.debug(CURRENT_ENTRY, entry);
            value = entry.getValue();
            key = entry.getKey();
            row = thirdSheet.getRow(value);
            if (secondAndThird.get(key) != null) {
                setAdditionValue(row, dailyThirdColumns,
                        withoutDied(secondAndThird.get(key).getChild()), // Выписано детей
                        secondAndThird.get(key).getChild().getDays(), // Проведено детей
                        secondAndThird.get(key).getChild().getDied()); // Умерло
                setAdditionValue(thirdSheet.getRow(8), dailyThirdColumns, // Сохранение во ВСЕГО
                        withoutDied(secondAndThird.get(key).getChild()),
                        secondAndThird.get(key).getChild().getDays(),
                        secondAndThird.get(key).getChild().getDied());
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
            Set<String> temp = new TreeSet<>(resultMapKeys);
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
     * сохранение значения в ячейку
     */
    private void setCellValue(Row row, List<Integer> columns, Integer... values) {
        for (int i = 0; i < columns.size(); i++) {
            row.getCell(columns.get(i)).setCellValue(values[i]);
        }
    }

    /**
     * сохранение суммы старого и нового значения в ячейку (для всего и "из них")
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
        Integer[] rowsFrom = {34, 33};
        Row rowTo = sheet.getRow(32);
        groupRows(rowTo, rowsFrom, sheet);
        // C81-C96 3.1.1
        rowsFrom = new Integer[]{26, 27, 28, 29, 30, 31, 32, 35, 36, 37, 38, 39};
        rowTo = sheet.getRow(25);
        groupRows(rowTo, rowsFrom, sheet);
        // C00-C97 3.1
        rowsFrom = new Integer[]{22, 23, 25, 40};
        rowTo = sheet.getRow(21);
        groupRows(rowTo, rowsFrom, sheet);

        // D10-D36 3.3
        rowsFrom = new Integer[]{42, 43, 44};
        rowTo = sheet.getRow(41);
        groupRows(rowTo, rowsFrom, sheet);

        // C00-D48 3.0
        rowsFrom = new Integer[]{21, 41, 44};
        rowTo = sheet.getRow(20);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupD5(Sheet sheet) {
        // D50-D64 4.1
        Integer[] rowsFrom = {47, 48};
        Row rowTo = sheet.getRow(46);
        groupRows(rowTo, rowsFrom, sheet);
        // D65-D69 4.2
        rowsFrom = new Integer[]{50, 51};
        rowTo = sheet.getRow(49);
        groupRows(rowTo, rowsFrom, sheet);
        // D50-D89 4.0
        rowsFrom = new Integer[]{46, 49, 52, 53};
        rowTo = sheet.getRow(45);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupE(Sheet sheet) {
        // E10-E14 5.4
        Integer[] rowsFrom = {59, 60, 61, 62, 63};
        Row rowTo = sheet.getRow(58);
        groupRows(rowTo, rowsFrom, sheet);
        // E00-E89 5.0
        rowsFrom = new Integer[]{55, 56, 57, 58, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76};
        rowTo = sheet.getRow(54);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupF(Sheet sheet) {
        // F01-F99 6.0
        Integer[] rowsFrom = {78, 79};
        Row rowTo = sheet.getRow(77);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupG(Sheet sheet) {
        // G00-G09 7.1
        Integer[] rowsFrom = {82, 83, 84};
        Row rowTo = sheet.getRow(81);
        groupRows(rowTo, rowsFrom, sheet);
        // G20-G25 7.3
        rowsFrom = new Integer[]{87, 88, 89};
        rowTo = sheet.getRow(86);
        groupRows(rowTo, rowsFrom, sheet);
        // G30-G31 7.4
        rowsFrom = new Integer[]{91, 92};
        rowTo = sheet.getRow(90);
        groupRows(rowTo, rowsFrom, sheet);
        // G35-G37 7.5
        rowsFrom = new Integer[]{94, 95};
        rowTo = sheet.getRow(93);
        groupRows(rowTo, rowsFrom, sheet);
        // G40-G47 7.6
        rowsFrom = new Integer[]{97, 98, 99};
        rowTo = sheet.getRow(96);
        groupRows(rowTo, rowsFrom, sheet);
        // G50-G64 7.7
        rowsFrom = new Integer[]{101, 102};
        rowTo = sheet.getRow(100);
        groupRows(rowTo, rowsFrom, sheet);
        // G70-G73 7.8
        rowsFrom = new Integer[]{104, 105, 106};
        rowTo = sheet.getRow(103);
        groupRows(rowTo, rowsFrom, sheet);
        // G80-G83 7.9
        rowsFrom = new Integer[]{108, 109};
        rowTo = sheet.getRow(107);
        groupRows(rowTo, rowsFrom, sheet);
        // G00-G98 7.0
        rowsFrom = new Integer[]{81, 85, 86, 90, 93, 96, 100, 103, 107, 110, 111, 112};
        rowTo = sheet.getRow(80);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupH(Sheet sheet) {
        // H46-H48 8.8
        Integer[] rowsFrom = {122, 123};
        Row rowTo = sheet.getRow(121);
        groupRows(rowTo, rowsFrom, sheet);
        // H54 8.9
        rowsFrom = new Integer[]{125, 126};
        rowTo = sheet.getRow(124);
        groupRows(rowTo, rowsFrom, sheet);
        // H00-H59 8.0
        rowsFrom = new Integer[]{114, 115, 116, 117, 118, 119, 120, 121, 124, 127};
        rowTo = sheet.getRow(113);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupH6(Sheet sheet) {
        // H65-H74 9.1
        Integer[] rowsFrom = {130, 131, 132, 133, 134, 135};
        Row rowTo = sheet.getRow(129);
        groupRows(rowTo, rowsFrom, sheet);
        // H81-H83 9.2
        rowsFrom = new Integer[]{137, 138, 139};
        rowTo = sheet.getRow(136);
        groupRows(rowTo, rowsFrom, sheet);
        // H90 9.3
        rowsFrom = new Integer[]{141, 142, 143};
        rowTo = sheet.getRow(140);
        groupRows(rowTo, rowsFrom, sheet);
        // H60-H95 9.0
        rowsFrom = new Integer[]{129, 136, 140, 144};
        rowTo = sheet.getRow(128);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupI(Sheet sheet) {
        // I05-I09 10.2
        Integer[] rowsFrom = {148, 149};
        Row rowTo = sheet.getRow(147);
        groupRows(rowTo, rowsFrom, sheet);
        // I10-I13 10.3
        rowsFrom = new Integer[]{151, 152, 153, 154};
        rowTo = sheet.getRow(150);
        groupRows(rowTo, rowsFrom, sheet);
        // I20 10.4.1
        rowsFrom = new Integer[]{157, 158};
        rowTo = sheet.getRow(156);
        groupRows(rowTo, rowsFrom, sheet);
        // I25 10.4.5
        rowsFrom = new Integer[]{163, 164};
        rowTo = sheet.getRow(162);
        groupRows(rowTo, rowsFrom, sheet);
        // I20-I25 10.4
        rowsFrom = new Integer[]{156, 159, 160, 161, 162};
        rowTo = sheet.getRow(155);
        groupRows(rowTo, rowsFrom, sheet);
        // I30-I51 10.6
        rowsFrom = new Integer[]{168, 169, 170, 171, 172, 173, 174, 175, 176, 177};
        rowTo = sheet.getRow(167);
        groupRows(rowTo, rowsFrom, sheet);
        // I67 10.7.6
        rowsFrom = new Integer[]{185, 186};
        rowTo = sheet.getRow(184);
        groupRows(rowTo, rowsFrom, sheet);
        // I60-I69 10.7
        rowsFrom = new Integer[]{179, 180, 181, 182, 183, 184, 187};
        rowTo = sheet.getRow(178);
        groupRows(rowTo, rowsFrom, sheet);
        // I80-I89 10.9
        rowsFrom = new Integer[]{190, 191, 192, 193};
        rowTo = sheet.getRow(189);
        groupRows(rowTo, rowsFrom, sheet);
        // I00-I99 10.0
        rowsFrom = new Integer[]{146, 147, 150, 155, 165, 166, 167, 178, 188, 189, 194};
        rowTo = sheet.getRow(145);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupJ(Sheet sheet) {
        // J00-J06 11.1
        Integer[] rowsFrom = {197, 198, 199};
        Row rowTo = sheet.getRow(196);
        groupRows(rowTo, rowsFrom, sheet);
        // J00-J98 11.0
        rowsFrom = new Integer[]{196, 200, 201, 202, 203, 204, 205, 206, 207, 208, 209, 210};
        rowTo = sheet.getRow(195);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupK(Sheet sheet) {
        // K50-K52 12.4
        Integer[] rowsFrom = {216, 217, 218};
        Row rowTo = sheet.getRow(215);
        groupRows(rowTo, rowsFrom, sheet);
        // K55-K63 12.5
        rowsFrom = new Integer[]{220, 221, 222, 223, 224, 225};
        rowTo = sheet.getRow(219);
        groupRows(rowTo, rowsFrom, sheet);
        // K70-K76 12.8
        rowsFrom = new Integer[]{228, 229};
        rowTo = sheet.getRow(227);
        groupRows(rowTo, rowsFrom, sheet);
        // K85-K86 12.10
        rowsFrom = new Integer[]{232, 233};
        rowTo = sheet.getRow(231);
        groupRows(rowTo, rowsFrom, sheet);
        // K00-K92 12.0
        rowsFrom = new Integer[]{212, 213, 214, 215, 219, 226, 227, 230, 231, 234};
        rowTo = sheet.getRow(207);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupL(Sheet sheet) {
        // L40 13.4
        Integer[] rowsFrom = {241, 242};
        Row rowTo = sheet.getRow(240);
        groupRows(rowTo, rowsFrom, sheet);
        // L00-L98 13.0
        rowsFrom = new Integer[]{236, 237, 238, 239, 240, 243, 244, 245};
        rowTo = sheet.getRow(235);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupM(Sheet sheet) {
        // M00-M25 14.1
        Integer[] rowsFrom = {248, 249, 250, 251, 252};
        Row rowTo = sheet.getRow(247);
        groupRows(rowTo, rowsFrom, sheet);
        // M30-M35 14.2
        rowsFrom = new Integer[]{254, 255};
        rowTo = sheet.getRow(253);
        groupRows(rowTo, rowsFrom, sheet);
        // M45-M49 14.4
        rowsFrom = new Integer[]{258, 259};
        rowTo = sheet.getRow(257);
        groupRows(rowTo, rowsFrom, sheet);
        // M80-M94 14.7
        rowsFrom = new Integer[]{263, 264};
        rowTo = sheet.getRow(262);
        groupRows(rowTo, rowsFrom, sheet);
        // M00-M99 14.0
        rowsFrom = new Integer[]{247, 253, 256, 257, 260, 261, 262, 265};
        rowTo = sheet.getRow(246);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupN(Sheet sheet) {
        // N70-N76 15.7
        Integer[] rowsFrom = {274, 275};
        Row rowTo = sheet.getRow(273);
        groupRows(rowTo, rowsFrom, sheet);
        // N00-N99 15.0
        rowsFrom = new Integer[]{267, 268, 269, 270, 271, 272, 273, 276, 277, 278, 279, 280};
        rowTo = sheet.getRow(266);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupQ(Sheet sheet) {
        // Q38-Q45 18.4
        Integer[] rowsFrom = {288, 289};
        Row rowTo = sheet.getRow(287);
        groupRows(rowTo, rowsFrom, sheet);
        // Q00-Q99 18.0
        rowsFrom = new Integer[]{284, 285, 286, 287, 290, 291, 292, 293, 294, 295};
        rowTo = sheet.getRow(283);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupR(Sheet sheet) {
        // R00-R99
        Integer[] rowsFrom = {306, 316, 331};
        Row rowTo = sheet.getRow(296);
        groupRows(rowTo, rowsFrom, sheet);
    }

    private void groupST(Sheet sheet) {
        // S*2-T14 20.1
        Integer[] rowsFrom = {348, 349};
        Row rowTo = sheet.getRow(347);
        groupRows(rowTo, rowsFrom, sheet);
        // T36-T50 20.5
        rowsFrom = new Integer[]{354, 355};
        rowTo = sheet.getRow(353);
        groupRows(rowTo, rowsFrom, sheet);
        // T51-T65 20.6
        rowsFrom = new Integer[]{357, 358};
        rowTo = sheet.getRow(356);
        groupRows(rowTo, rowsFrom, sheet);
        // Soo-T98 20.0
        rowsFrom = new Integer[]{347, 350, 351, 352, 353, 356, 359};
        rowTo = sheet.getRow(346);
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
        Row from = sheet.getRow(9); // ВСЕГО
        Row z = sheet.getRow(361); // Z00-Z99 21.0
        for (Integer column : fourteenColumns) {
            double valueFrom = from.getCell(column).getNumericCellValue();
            double valueZ = z.getCell(column).getNumericCellValue();
            from.getCell(column).setCellValue(valueFrom - valueZ);
        }
    }
}