package com.example.reports.service;

import com.example.reports.entity.Data;
import com.example.reports.entity.People;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;

import java.io.File;
import java.io.FileInputStream;
import java.nio.file.Path;
import java.util.Map;

public interface ReportService {

    void process() throws Exception;

    void toTempFile(String folder) throws Exception;

    Map<String, Data> getResult();

    default Sheet getSheet(Path from) throws Exception {
        // TODO: 19.11.2022 считать excel, вернуть sheet
        try (FileInputStream file = new FileInputStream(from.toFile())) {
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            return workbook.getSheetAt(0);
        }
    }

    default void check(Map<String, Data> result, Logger log) {
        result.forEach((key, value) -> {
            People all = value.getAll();
            People adult = value.getAdult();
            People child = value.getChild();
            People old = value.getOld();
            checkEquals(all.getAll(), adult.getAll(), child.getAll(), old.getAll(), "getAll", log);
            checkEquals(all.getAmbulance(), adult.getAmbulance(), child.getAmbulance(), old.getAmbulance(), "getAmbulance", log);
            checkEquals(all.getDays(), adult.getDays(), child.getDays(), old.getDays(), "getDays", log);
            checkEquals(all.getDied(), adult.getDied(), child.getDied(), old.getDied(), "getDied", log);
            checkEquals(all.getEmergency(), adult.getEmergency(), child.getEmergency(), old.getEmergency(), "getEmergency", log);
        });
        log.debug("Result: {}", result);
    }

    default void checkEquals(int all, int adult, int child, int old, String method, Logger log) {
        if (all != (adult + child + old)) {
            log.error("{}() not equals", method);
        }
    }
}
