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

import static com.example.reports.util.PathsConstants.FOR_FIRST;
import static com.example.reports.util.PathsConstants.FOR_FIRST_POST;

// Works with baza and postup
@Service
@RequiredArgsConstructor
@Slf4j
public class FirstFormPostService implements ReportService {

    private final RowToDataMapper mapper;
    private final HelperForTemp helper;
    private final AppConfig appConfig;

    private List<Integer> fromColumns;
    private Map<String, Data> result;

    @PostConstruct
    void init() {
        fromColumns = appConfig.getFrom().getPost();
    }

    @Override
    public void process() throws Exception {
        log.info("start process of {}", FirstFormPostService.class.getName());
        Sheet sheet = getSheet(FOR_FIRST_POST);
        log.debug("Sheet name: " + sheet.getSheetName());
        Iterator<Row> rowIterator = sheet.rowIterator();
        Row skipHeader = rowIterator.next();
        log.trace(skipHeader.toString());
        while (rowIterator.hasNext()) {
            Row next = rowIterator.next();
            if (!next.getCell(fromColumns.get(0)).getStringCellValue().toLowerCase().contains("(дневной стационар)")) {
                mapper.mapPost(next, fromColumns);
            }
        }
        result = mapper.getResult();
        check(result, log);
        toTempFile("post");
        log.info("end process of {}", FirstFormPostService.class.getName());
        FOR_FIRST_POST.toFile().deleteOnExit();
    }

    @Override
    public void toTempFile(String folder) throws Exception {
        List<String> strings = new ArrayList<>();
        strings.add(
                "Койки\t" +
                "Всего\t" +
                "Село\t" +
                "Дети\t" +
                "Взрослые\t" +
                "Пенсионеры\n");
        result.forEach((key, value) -> strings.add(
                key + "\t" +
                value.getAll().getAll() + "\t" +
                value.getVillage() + "\t" +
                value.getChild().getAll() + "\t" +
                value.getAdult().getAll() + "\t" +
                value.getOld().getAll() + "\n"));
        Path path = FOR_FIRST.resolve(Path.of(folder));
        Files.createDirectories(path);
        helper.writeFile(path.resolve(PathsConstants.TEMP), strings);
    }

    @Override
    public Map<String, Data> getResult() {
        return result;
    }
}
