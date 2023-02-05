package com.example.reports.service;

import com.example.reports.config.AppConfig;
import com.example.reports.entity.Data;
import com.example.reports.entity.SecondFormPeople;
import com.example.reports.mapper.RowToDataMapper;
import com.example.reports.util.PathsConstants;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.math3.util.Pair;
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

import static com.example.reports.util.PathsConstants.FOR_SECOND;
import static com.example.reports.util.PathsConstants.FOR_SECOND_BAZA;

//works with baza
@Service
@RequiredArgsConstructor
@Slf4j
public class SecondFormService implements ReportService {

    private final RowToDataMapper mapper;
    private final HelperForTemp helper;
    private final AppConfig appConfig;

    private List<Integer> fromColumns;
    private Map<String, Pair<SecondFormPeople, SecondFormPeople>> result;

    @PostConstruct
    void init() {
        fromColumns = appConfig.getFrom().getSecondFormBaza();
    }

    @Override
    public void process() throws Exception {
        log.info("start process of {}", SecondFormService.class.getName());
        Sheet sheet = getSheet(FOR_SECOND_BAZA);
        log.debug("Sheet name: " + sheet.getSheetName());
        Iterator<Row> rowIterator = sheet.rowIterator();
        Row skipHeader = rowIterator.next();
        log.trace(skipHeader.toString());
        while (rowIterator.hasNext()) {
            Row next = rowIterator.next();
            if (!next.getCell(fromColumns.get(0)).getStringCellValue().toLowerCase().contains("(дневной стационар)")) {
                mapper.mapBazaForSecondForm(next, fromColumns);
            }
        }
        result = mapper.getSecondFormResult();
        toTempFile("temp");
        log.info("end process of {}", SecondFormService.class.getName());
        FOR_SECOND_BAZA.toFile().deleteOnExit();
    }

    @Override
    public void toTempFile(String folder) throws Exception {
        List<String> strings = new ArrayList<>();
        strings.add(
                "МКБ\t" +
                "0-14 лет\t" +
                "15-19 лет\t" +
                "20-24 года\t" +
                "25-29 лет\t" +
                "30-34 года\t" +
                "35-39 лет\t" +
                "40-44 года\t" +
                "45-49 лет\t" +
                "50-54 года\t" +
                "55-59 лет\t" +
                "60-64 года\t" +
                "65-69 лет\t" +
                "70-74 года\t" +
                "75-79 лет\t" +
                "80-84 года\t" +
                "85 лет и старше\t" +
                "Умерло 0-14 лет\t" +
                "Умерло 15-19 лет\t" +
                "Умерло 20-24 года\t" +
                "Умерло 25-29 лет\t" +
                "Умерло 30-34 года\t" +
                "Умерло 35-39 лет\t" +
                "Умерло 40-44 года\t" +
                "Умерло 45-49 лет\t" +
                "Умерло 50-54 года\t" +
                "Умерло 55-59 лет\t" +
                "Умерло 60-64 года\t" +
                "Умерло 65-69 лет\t" +
                "Умерло 70-74 года\t" +
                "Умерло 75-79 лет\t" +
                "Умерло 80-84 года\t" +
                "Умерло 85 лет и старше\n");
        result.forEach((key, value) -> strings.add(
                key + "\t" +
                value.getFirst().getBelow14() + "\t" +
                value.getFirst().getBetween15_19() + "\t" +
                value.getFirst().getBetween20_24() + "\t" +
                value.getFirst().getBetween25_29() + "\t" +
                value.getFirst().getBetween30_34() + "\t" +
                value.getFirst().getBetween35_39() + "\t" +
                value.getFirst().getBetween40_44() + "\t" +
                value.getFirst().getBetween45_49() + "\t" +
                value.getFirst().getBetween50_54() + "\t" +
                value.getFirst().getBetween55_59() + "\t" +
                value.getFirst().getBetween60_64() + "\t" +
                value.getFirst().getBetween65_69() + "\t" +
                value.getFirst().getBetween70_74() + "\t" +
                value.getFirst().getBetween75_79() + "\t" +
                value.getFirst().getBetween80_84() + "\t" +
                value.getFirst().getAbove85() + "\t" +
                value.getSecond().getBelow14() + "\t" +
                value.getSecond().getBetween15_19() + "\t" +
                value.getSecond().getBetween20_24() + "\t" +
                value.getSecond().getBetween25_29() + "\t" +
                value.getSecond().getBetween30_34() + "\t" +
                value.getSecond().getBetween35_39() + "\t" +
                value.getSecond().getBetween40_44() + "\t" +
                value.getSecond().getBetween45_49() + "\t" +
                value.getSecond().getBetween50_54() + "\t" +
                value.getSecond().getBetween55_59() + "\t" +
                value.getSecond().getBetween60_64() + "\t" +
                value.getSecond().getBetween65_69() + "\t" +
                value.getSecond().getBetween70_74() + "\t" +
                value.getSecond().getBetween75_79() + "\t" +
                value.getSecond().getBetween80_84() + "\t" +
                value.getSecond().getAbove85() + "\n"));
        Path path = FOR_SECOND.resolve(Path.of(folder));
        Files.createDirectories(path);
        helper.writeFile(path.resolve(PathsConstants.TEMP), strings);
    }

    @Override
    public Map<String, Data> getResult() {
        throw new UnsupportedOperationException("method not implemented");
    }

    public Map<String, Pair<SecondFormPeople, SecondFormPeople>> getSecondFormResult() {
        return result;
    }

    // Напишу пока тут: берем из базы 3 столбца МКБ, возраст, результат
    // далее по МКБ выбираем то, что нужно для 2910 (записываем в tempFile и соответственно в result), - это в маппере
    // если нашли нужный МКБ, смотрим результат госпитализации
    // если умер, то смотрим возраст и распределяем по столбцу,
    // если не умер, то смотрим возраст и распределяем по столбцу - это в ReportSaving
}
