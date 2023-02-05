package com.example.reports.service;

import com.example.reports.config.AppConfig;
import com.example.reports.entity.Data;
import com.example.reports.mapper.RowToDataMapper;
import com.example.reports.util.PathsConstants;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.stereotype.Service;

import javax.annotation.PostConstruct;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import static com.example.reports.util.PathsConstants.FOR_DAILY;
import static com.example.reports.util.PathsConstants.FOR_DAILY_SECOND;

//works with baza
@Service
@RequiredArgsConstructor
@Slf4j
public class DailyFormSecondSheetService implements ReportService {

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
        log.info("start process of {}", DailyFormSecondSheetService.class.getName());
        Sheet sheet = getSheet(FOR_DAILY_SECOND);
        log.debug("Sheet name: " + sheet.getSheetName());
        Iterator<Row> rowIterator = sheet.rowIterator();
        Row skipHeader = rowIterator.next();
        log.trace(skipHeader.toString());
        while (rowIterator.hasNext()) {
            Row next = rowIterator.next();
            if (next.getCell(fromColumns.get(0)).getStringCellValue().toLowerCase().contains("(дневной стационар)")) {
                mapper.mapBazaByMKB(next, fromColumns);
            }
        }
        result = mapper.getResult();
        check(result, log);
        toTempFile("3000");
        log.info("end process of {}", DailyFormSecondSheetService.class.getName());
        FOR_DAILY_SECOND.toFile().deleteOnExit();
    }

    @Override
    public void toTempFile(String folder) throws Exception {
        List<String> strings = new ArrayList<>();
        strings.add(
                "МКБ\t" +
                "Дети\t" +
                "Проведено(дет)\t" +
                "Умерло(дет)\t" +
                "Взрослые\t" +
                "Проведено(взр)\t" +
                "Умерло(взр)\n");
        result.forEach((key, value) -> strings.add(
                key + "\t" +
                (value.getChild().getAll() - value.getChild().getDied()) + "\t" +
                value.getChild().getDays() + "\t" +
                value.getChild().getDied() + "\t" +
                ((value.getAdult().getAll() - value.getAdult().getDied()) + (value.getOld().getAll() - value.getOld().getDied())) + "\t" +
                (value.getAdult().getDays() + value.getOld().getDays()) + "\t" +
                (value.getAdult().getDied() + value.getOld().getDied()) + "\n"));
        Path path = FOR_DAILY.resolve(Path.of(folder));
        Files.createDirectories(path);
        helper.writeFile(path.resolve(PathsConstants.TEMP), strings);
    }

    @Override
    public Map<String, Data> getResult() {
        return result;
    }
}
