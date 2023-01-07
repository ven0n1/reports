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
import static com.example.reports.util.PathsConstants.FOR_FOURTEEN;
import static com.example.reports.util.PathsConstants.FOR_FOURTEEN_BAZA;

//works with baza
@Service
@RequiredArgsConstructor
@Slf4j
public class FourteenFormService implements ReportService {

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
        log.info("start process of {}", FourteenFormService.class.getName());
        // TODO: 19.11.2022 открыть excel из origin/for_fourteen считывать строки, вызывать маппер
        Sheet sheet = getSheet(FOR_FOURTEEN_BAZA);
        log.debug("Sheet name: " + sheet.getSheetName());
        Iterator<Row> rowIterator = sheet.rowIterator();
        Row skipHeader = rowIterator.next();
        log.trace(skipHeader.toString());
        while (rowIterator.hasNext()) {
            Row next = rowIterator.next();
            if (!next.getCell(fromColumns.get(0)).getStringCellValue().toLowerCase().contains("(дневной стационар)")) {
                mapper.mapBazaByMKB(next, fromColumns);
            }
        }
        result = mapper.getResult();
        check(result, log);
        toTempFile("temp");
        log.info("end process of {}", FourteenFormService.class.getName());
        FOR_FOURTEEN_BAZA.toFile().deleteOnExit();
    }

    @Override
    public void toTempFile(String folder) throws Exception {
        List<String> strings = new ArrayList<>();
        strings.add(
                "МКБ\t" +
                "Выписано(дет)\t" +
                "По экстренным(дет)\t" +
                "На скорой(дет)\t" +
                "Койко-дней(дет)\t" +
                "Умерло(дет)\t" +
                "Выписано(взр)\t" +
                "По экстренным(взр)\t" +
                "На скорой(взр)\t" +
                "Койко-дней(взр)\t" +
                "Умерло(взр)\t" +
                "Выписано(пен)\t" +
                "По экстренным(пен)\t" +
                "На скорой(пен)\t" +
                "Койко-дней(пен)\t" +
                "Умерло(пен)\n");
        result.forEach((key, value) -> strings.add(
                key + "\t" +
                (value.getChild().getAll() - value.getChild().getDied()) + "\t" +
                value.getChild().getEmergency() + "\t" +
                value.getChild().getAmbulance() + "\t" +
                value.getChild().getDays() + "\t" +
                value.getChild().getDied() + "\t" +
                (value.getAdult().getAll() - value.getAdult().getDied()) + "\t" +
                value.getAdult().getEmergency() + "\t" +
                value.getAdult().getAmbulance() + "\t" +
                value.getAdult().getDays() + "\t" +
                value.getAdult().getDied() + "\t" +
                (value.getOld().getAll() - value.getOld().getDied()) + "\t" +
                value.getOld().getEmergency() + "\t" +
                value.getOld().getAmbulance() + "\t" +
                value.getOld().getDays() + "\t" +
                value.getOld().getDied() + "\n"));
        Path path = FOR_FOURTEEN.resolve(Path.of(folder));
        Files.createDirectories(path);
        helper.writeFile(path.resolve(PathsConstants.TEMP), strings);
    }

    @Override
    public Map<String, Data> getResult() {
        return result;
    }
}
