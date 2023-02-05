package com.example.reports.mapper;

import com.example.reports.config.AppConfig;
import com.example.reports.entity.Data;
import com.example.reports.entity.People;
import com.example.reports.entity.SecondFormPeople;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.math3.util.Pair;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.beans.factory.config.ConfigurableBeanFactory;
import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Component;

import javax.annotation.PostConstruct;
import java.util.EnumMap;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.function.Function;
import java.util.function.Supplier;

@Component
@Scope(value = ConfigurableBeanFactory.SCOPE_PROTOTYPE)
@RequiredArgsConstructor
@Slf4j
public class RowToDataMapper {

    private final Map<String, Data> result = new HashMap<>();
    private Map<String, Pair<SecondFormPeople, SecondFormPeople>> secondFormResult = new HashMap<>();
    private final AppConfig appConfig;

    private int ageWoman;
    private int ageMan;
    private Map<String, List<String>> secondFormMkbCodes;

    private final Map<CellType, Function<Cell, String>> cellToString = new EnumMap<>(CellType.class);
    private String profile;
    private String gender;
    private String age;
    private String region;
    private String duration;
    private String resultOf;
    private String mkb;
    private String delivered;
    private String urgent;
    private String ageCategory;
    public static final String ROW_NUM = "Row num: {}";

    @PostConstruct
    void init() {
        cellToString.put(CellType.STRING, Cell::getStringCellValue);
        cellToString.put(CellType.NUMERIC, cell -> String.valueOf(cell.getNumericCellValue()));
        cellToString.put(CellType.BLANK, cell -> "");
        cellToString.put(CellType.FORMULA, cell -> String.valueOf(cellToString.get(cell.getCachedFormulaResultType()).apply(cell)));
        ageWoman = appConfig.getAge().get("woman");
        ageMan = appConfig.getAge().get("man");
        secondFormMkbCodes = appConfig.getSecondFormMkbCodes();
    }

    public void mapBazaByProfile(Row row, List<Integer> fromColumns) {
        log.trace(ROW_NUM, row.getRowNum());
        setVariable(row, fromColumns);
        mapRow(row, profile);
    }

    public void mapBazaByMKB(Row row, List<Integer> fromColumns) {
        log.trace(ROW_NUM, row.getRowNum());
        setVariable(row, fromColumns);
        mapRow(row, mkb);
    }

    public Map<String, Data> getResult() {
        return result;
    }

    public void mapPost(Row row, List<Integer> fromColumns) {
        log.trace(ROW_NUM, row.getRowNum());
        gender = row.getCell(fromColumns.get(1)).getStringCellValue(); // пол
        age = convertToString(row.getCell(fromColumns.get(2))); // возраст
        region = wrapException(() -> row.getCell(fromColumns.get(3)).getStringCellValue()); // город/село
        profile = row.getCell(fromColumns.get(4)).getStringCellValue(); // профиль койки
        String formatedKey = profile.toLowerCase().strip();
        result.putIfAbsent(formatedKey, new Data());
        Data data = result.get(formatedKey);
        People all = data.getAll();
        People child = data.getChild();
        People adult = data.getAdult();
        People old = data.getOld();
        all.setAll(setAll(all.getAll()));
        data.setVillage(setVillage(region, data.getVillage()));
        try {
            double parsedAge = Double.parseDouble(age);
            if (parsedAge < 18.0) {
                log.trace("Child " + ROW_NUM, row.getRowNum());
                child.setAll(setAll(child.getAll()));
            } else if ((gender.equalsIgnoreCase("муж") && parsedAge < ageMan) || (gender.equalsIgnoreCase("жен") && parsedAge < ageWoman)) {
                log.trace("Adult " + ROW_NUM, row.getRowNum());
                adult.setAll(setAll(adult.getAll()));
            } else {
                log.trace("Aged " + ROW_NUM, row.getRowNum());
                old.setAll(setAll(old.getAll()));
            }
        } catch (Exception e) {
            log.info("Adult " + ROW_NUM, row.getRowNum());
            adult.setAll(setAll(adult.getAll()));
        }
    }

    public void mapBazaForSecondForm(Row row, List<Integer> fromColumns) {
        log.trace(ROW_NUM, row.getRowNum());
        mkb = convertToString(row.getCell(fromColumns.get(1))).toLowerCase(Locale.ROOT).strip(); // МКБ заключительный
        age = convertToString(row.getCell(fromColumns.get(2))); // возраст
        resultOf = convertToString(row.getCell(fromColumns.get(3))); // результат госпитализации
        secondFormMkbCodes.forEach((key, value) -> {
            secondFormResult.putIfAbsent(key, new Pair<>(new SecondFormPeople(), new SecondFormPeople()));
            value.forEach(code -> {
                if (mkb.startsWith(code)) {
                    if (resultOf.equalsIgnoreCase("умер")) {
                        SecondFormPeople dead = secondFormResult.get(key).getSecond();
                        parseAge(dead);
                    } else {
                        SecondFormPeople alive = secondFormResult.get(key).getFirst();
                        parseAge(alive);
                    }
                }
            });
        });
    }

    private void parseAge(SecondFormPeople people) {
        try {
            double parsedAge = Double.parseDouble(age);
            log.trace("age {}", parsedAge);
            mapSecondFormPeopleByAge(people, parsedAge);
        } catch (NumberFormatException e) {
            double parsedAge = Double.parseDouble(age.strip().substring(0, age.length() - 5).strip());
            log.info("age {}", parsedAge);
            mapSecondFormPeopleByAge(people, parsedAge);
        }
    }

    public Map<String, Pair<SecondFormPeople, SecondFormPeople>> getSecondFormResult() {
        return secondFormResult;
    }

    private void mapRow(Row row, String key) {
        String formatedKey = key.toLowerCase().strip();
        result.putIfAbsent(formatedKey, new Data());
        Data data = result.get(formatedKey);
        data.setAll(setPeople(data.getAll(), urgent, delivered, duration, resultOf));
        distributeByAge(row, age, duration, resultOf, delivered, urgent, ageCategory, data);}

    private void setVariable(Row row, List<Integer> fromColumns) {
        gender = convertToString(row.getCell(fromColumns.get(1))); // пол
        age = convertToString(row.getCell(fromColumns.get(2))); // возраст
        duration = convertToString(row.getCell(fromColumns.get(3))); // продолжительность госпитализации
        resultOf = convertToString(row.getCell(fromColumns.get(4))); // результат госпитализации
        mkb = convertToString(row.getCell(fromColumns.get(5))); // МКБ заключительный
        profile = convertToString(row.getCell(fromColumns.get(6))); // профиль койки
        delivered = convertToString(row.getCell(fromColumns.get(7))); // кем доставлен
        urgent = convertToString(row.getCell(fromColumns.get(8))); // плановая/экстренная
        ageCategory = convertToString(row.getCell(fromColumns.get(9))); // для 14 кс (возрастная категория)
    }

    private void distributeByAge(Row row, String age, String duration, String resultOf, String delivered, String urgent, String ageCategory, Data data) {
//        try {
//            double parsedAge = Double.parseDouble(age);
//            log.trace("age: {}", parsedAge);
//            if (ageCategory.toLowerCase().endsWith("труд")) {
//                log.trace("Adult " + ROW_NUM, row.getRowNum());
//                data.setAdult(setPeople(data.getAdult(), urgent, delivered, duration, resultOf));
//            } else if (ageCategory.toLowerCase().endsWith("дети")) {
//                log.trace("Child " + ROW_NUM, row.getRowNum());
//                data.setChild(setPeople(data.getChild(), urgent, delivered, duration, resultOf));
//            } else if (ageCategory.toLowerCase().endsWith("пенсион")) {
//                log.trace("Aged " + ROW_NUM, row.getRowNum());
//                data.setOld(setPeople(data.getOld(), urgent, delivered, duration, resultOf));
//            } else {
//                log.error("Строка: {} не распределена", (row.getRowNum() + 1));
//            }
//        } catch (NumberFormatException e) {
//            log.info("Adult " + ROW_NUM, row.getRowNum());
//            data.setAdult(setPeople(data.getAdult(), urgent, delivered, duration, resultOf));
//        }

        try {
            double parsedAge = Double.parseDouble(age);
            if (parsedAge < 18.0) {
                log.trace("Child " + ROW_NUM, row.getRowNum());
                data.setChild(setPeople(data.getChild(), urgent, delivered, duration, resultOf));
            } else if ((gender.equalsIgnoreCase("муж") && parsedAge < ageMan) || (gender.equalsIgnoreCase("жен") && parsedAge < ageWoman)) {
                log.trace("Adult " + ROW_NUM, row.getRowNum());
                data.setAdult(setPeople(data.getAdult(), urgent, delivered, duration, resultOf));
            } else {
                log.trace("Aged " + ROW_NUM, row.getRowNum());
                data.setOld(setPeople(data.getOld(), urgent, delivered, duration, resultOf));
            }
        } catch (Exception e) {
            log.info("Adult " + ROW_NUM, row.getRowNum());
            data.setAdult(setPeople(data.getAdult(), urgent, delivered, duration, resultOf));
        }
    }

    private void mapSecondFormPeopleByAge(SecondFormPeople people, double age) {
        if (age < 15.0) {
            people.incBelow14();
        } else if (age < 20.0) {
            people.incBetween15_19();
        } else if (age < 25.0) {
            people.incBetween20_24();
        } else if (age < 30.0) {
            people.incBetween25_29();
        } else if (age < 35.0) {
            people.incBetween30_34();
        } else if (age < 40.0) {
            people.incBetween35_39();
        } else if (age < 45.0) {
            people.incBetween40_44();
        } else if (age < 50.0) {
            people.incBetween45_49();
        } else if (age < 55.0) {
            people.incBetween50_54();
        } else if (age < 60.0) {
            people.incBetween55_59();
        } else if (age < 65.0) {
            people.incBetween60_64();
        } else if (age < 70.0) {
            people.incBetween65_69();
        } else if (age < 75.0) {
            people.incBetween70_74();
        } else if (age < 80.0) {
            people.incBetween75_79();
        } else if (age < 85.0) {
            people.incBetween80_84();
        } else {
            people.incAbove85();
        }
    }

    private String wrapException(Supplier<String> supplier) {
        String supplierResult = "";
        try {
            supplierResult = supplier.get();
        } catch (Exception e) {
            //skip
        }
        return supplierResult;
    }

    private String convertToString(Cell cell) {
        return cellToString.get(cell.getCellType()).apply(cell);
    }

    private People setPeople(People people, String urgent, String delivered, String duration, String resultOf) {
        people.setAll(setAll(people.getAll()));
        people.setEmergency(setEmergency(urgent, people.getEmergency()));
        people.setAmbulance(setAmbulance(delivered, people.getAmbulance()));
        people.setDays(setDays(duration, people.getDays()));
        people.setDied(setDied(resultOf, people.getDied()));
        return people;
    }

    private int setAll(int all) {
        return ++all;
    }

    private int setEmergency(String urgent, int emergency) {
//        if (urgent.equalsIgnoreCase("экстренная")) {
//            emergency++;
//        }
        if (urgent.equalsIgnoreCase("Доставлен бригадой скорой помощи") ||
            urgent.startsWith("МБУЗ") || urgent.equalsIgnoreCase("Экстренно")) {
            emergency++;
        }
        return emergency;
    }

    private int setAmbulance(String delivered, int ambulance) {
//        if (delivered.equalsIgnoreCase("скорая помощь")) {
//            ambulance++;
//        }
        if (delivered.equalsIgnoreCase("Доставлен бригадой скорой помощи") ||
                delivered.startsWith("МБУЗ") || delivered.equalsIgnoreCase("Экстренно")) {
            ambulance++;
        }
        return ambulance;
    }

    private int setDays(String duration, int alreadyDuration) {
        double parsed = Double.parseDouble(duration);
        alreadyDuration += parsed;
        return alreadyDuration;
    }

    private int setDied(String resultOf, int died) {
        if (resultOf.equalsIgnoreCase("умер")) {
            died++;
        }
        return died;
    }

    private int setVillage(String region, int village) {
        if (region.equalsIgnoreCase("село")) {
            village++;
        }
        return village;
    }
}
