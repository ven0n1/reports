package com.example.reports.service;

import com.example.reports.config.AppConfig;
import com.example.reports.entity.Data;
import com.example.reports.mapper.RowToDataMapper;
import com.example.reports.util.PathsConstants;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import javax.annotation.PostConstruct;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import static com.example.reports.util.PathsConstants.BAZA;
import static com.example.reports.util.PathsConstants.FOR_DAILY;
import static com.example.reports.util.PathsConstants.FOR_DAILY_FIRST;
import static com.example.reports.util.PathsConstants.FOR_FIRST;
import static com.example.reports.util.PathsConstants.FOR_FIRST_BAZA;

// Works with baza and postup
@Service
@RequiredArgsConstructor
@Slf4j
public class FirstFormBazaService implements ReportService {

    private final RowToDataMapper mapper;
    private final HelperForTemp helper;
    private final AppConfig appConfig;

    private List<Integer> fromColumns;
    private Map<String, Data> result;

    @PostConstruct
    void init() {
        fromColumns = appConfig.getFrom().getBaza();
    }

    @Override
    public void process() throws Exception {
        log.info("start process of {}", FirstFormBazaService.class.getName());
        // TODO: 19.11.2022 открыть excel из origin/for_first считывать строки, вызывать маппер
        Sheet sheet = getSheet(FOR_FIRST_BAZA);
        log.debug("Sheet name: " + sheet.getSheetName());
        Iterator<Row> rowIterator = sheet.rowIterator();
        Row skipHeader = rowIterator.next();
        log.trace(skipHeader.toString());
        while (rowIterator.hasNext()) {
            Row next = rowIterator.next();
            if (!next.getCell(fromColumns.get(0)).getStringCellValue().toLowerCase().contains("(дневной стационар)")) {
                mapper.mapBazaByProfile(next, fromColumns);
            }
        }
        result = mapper.getResult();
        check(result, log);
        toTempFile("baza");
        log.info("end process of {}", FirstFormBazaService.class.getName());
        FOR_FIRST_BAZA.toFile().deleteOnExit();
    }

    @Override
    public void toTempFile(String folder) throws Exception {
        List<String> strings = new ArrayList<>();
        strings.add(
                "Койки\t" +
                "Всего\t" +
                "Проведено(всего)\t" +
                "Умерло(всего)\t" +
                "Дети\t" +
                "Проведено(дет)\t" +
                "Умерло(дет)\t" +
                "Взрослые\t" +
                "Проведено(взр)\t" +
                "Умерло(взр)\t" +
                "Пенсионеры\t" +
                "Проведено(пен)\t" +
                "Умерло(пен)\n");
        result.forEach((key, value) -> strings.add(
                key + "\t" +
                (value.getAll().getAll() - value.getAll().getDied()) + "\t" +
                value.getAll().getDays() + "\t" +
                value.getAll().getDied() + "\t" +
                (value.getChild().getAll() - value.getChild().getDied()) + "\t" +
                value.getChild().getDays() + "\t" +
                value.getChild().getDied() + "\t" +
                (value.getAdult().getAll() - value.getAdult().getDied()) + "\t" +
                value.getAdult().getDays() + "\t" +
                value.getAdult().getDied() + "\t" +
                (value.getOld().getAll() - value.getOld().getDied()) + "\t" +
                value.getOld().getDays() + "\t" +
                value.getOld().getDied() + "\n"));
        Path path = FOR_FIRST.resolve(Path.of(folder));
        Files.createDirectories(path);
        helper.writeFile(path.resolve(PathsConstants.TEMP), strings);
    }

    @Override
    public Map<String, Data> getResult() {
        return result;
    }
}
