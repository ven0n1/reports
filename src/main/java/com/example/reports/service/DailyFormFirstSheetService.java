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

//works with baza
@Service
@RequiredArgsConstructor
@Slf4j
public class DailyFormFirstSheetService implements ReportService {

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
    public void process() throws Exception{
        log.info("start process of {}", DailyFormFirstSheetService.class.getName());
        // TODO: 19.11.2022 открыть excel из origin/for_daily считывать строки, вызывать маппер
        Sheet sheet = getSheet(FOR_DAILY_FIRST);
        log.debug("Sheet name: " + sheet.getSheetName());
        Iterator<Row> rowIterator = sheet.rowIterator();
        Row skipHeader = rowIterator.next();
        log.trace(skipHeader.toString());
        while (rowIterator.hasNext()) {
            Row next = rowIterator.next();
            if (next.getCell(fromColumns.get(0)).getStringCellValue().toLowerCase().contains("(дневной стационар)")) {
                mapper.mapBazaByProfile(next, fromColumns);
            }
        }
        result = mapper.getResult();
        check(result, log);
        toTempFile("2000");
        log.info("end process of {}", DailyFormFirstSheetService.class.getName());
        FOR_DAILY_FIRST.toFile().deleteOnExit();
    }

    @Override
    public void toTempFile(String folder) throws Exception {
        List<String> strings = new ArrayList<>();
        strings.add(
                "Койки\t" +
                "Взрослые\t" +
                "Пенсионеры\t" +
                "Дети\t" +
                "Проведено(взр)\t" +
                "Проведено(пен)\t" +
                "Проведено(дети)\n");
        result.forEach((key, value) -> strings.add(
                key + "\t" +
                value.getAdult().getAll() + "\t" +
                value.getOld().getAll() + "\t" +
                value.getChild().getAll() + "\t" +
                value.getAdult().getDays() + "\t" +
                value.getOld().getDays() + "\t" +
                value.getChild().getDays() + "\n"));
        Path path = FOR_DAILY.resolve(Path.of(folder));
        Files.createDirectories(path);
        helper.writeFile(path.resolve(PathsConstants.TEMP), strings);
    }

    @Override
    public Map<String, Data> getResult() {
        return result;
    }
}
