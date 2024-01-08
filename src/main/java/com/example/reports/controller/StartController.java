package com.example.reports.controller;

import com.example.reports.service.FourteenFormService;
import com.example.reports.service.ReportService;
import com.example.reports.service.ReportsSaving;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FileUtils;
import org.springframework.web.bind.annotation.RestController;

import javax.annotation.PostConstruct;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static com.example.reports.util.PathsConstants.FOR_FOURTEEN_BAZA;
import static com.example.reports.util.PathsConstants.FROM_BAZA;

@RestController
@Slf4j
@RequiredArgsConstructor
public class StartController {

    private final List<ReportService> reports;
    private final ReportsSaving reportsSaving;

    @PostConstruct
    public void init() throws Exception {

//        copyToDirectory(FOR_DAILY_FIRST, FROM_BAZA);
//        copyToDirectory(FOR_DAILY_SECOND, FROM_BAZA);
//        checkSpecificService(DailyFormFirstSheetService.class);
//        checkSpecificService(DailyFormSecondSheetService.class);
//        reportsSaving.saveToDailyForm();

//        copyToDirectory(FOR_FIRST_BAZA, FROM_BAZA);
//        copyToDirectory(FOR_FIRST_POST, FROM_POST);
//        checkSpecificService(FirstFormBazaService.class);
//        checkSpecificService(FirstFormPostService.class);
//        reportsSaving.saveToFirstForm();

        copyToDirectory(FOR_FOURTEEN_BAZA, FROM_BAZA);
        checkSpecificService(FourteenFormService.class);
        reportsSaving.saveToFourteenForm();

//        copyToDirectory(FOR_SECOND_BAZA, FROM_BAZA);
//        checkSpecificService(SecondFormService.class);
//        reportsSaving.saveToSecondForm(); // в столбец: всего (5б 23) нужно самостоятельно закинуть (лень это в коде прописывать)
    }

//    public static void main(String[] args) {
//        String s = "m33.2, m33.1, m79.3, m79.5, c16.8, i72.3, k43, c80.9, i72.8, n31.9, k50, r16.1, i15, c91.5, g54.8, i50.1, i08.3, f32.0, d47.4, d47.3, f32.8, f32.2, m31.7, c18.2, d80.0, m31.3, e25.0, d16.1, f23.2, i42, i74.2, c82.7, i48, c82.1, c82.4, c93.0, g43.8, s69.7, m99.0, f45.3, f45.0, g21.1, s12.2, t87.6, m75.4, e27.9, m75.0, s43.0, h43.8, d29.3, c84.6, e16.1, c84.1, l53.1, g58.0, g12.1, m51.9, c95.0, i65.3, e03.1, g12.9, i65.8, k61.0, g58.7, l53.9, n82.1, k85.9, d73.9, g36.9, l40.3, m84.0, m84.4, i67.3, k63.8, i10.0, k63.1, k63.0, c62.9, g25.2, c62.1, n04.0, n04.3, n04.5, c44, t95.1, n39.0, i22.1, l08.9, t84.3, e83.1, m17.5, m17.4, i35.0, j33.0, k31.8, c41.0, g04.2, k31.3, h02.0, g70.1, k66.9, n48.9, c82, l94.9, c67.8, q64.1, c67.5, t93.5, c67.1, f50.2, c94, c97, i48.2, i48.4, d30.0, d30.4, n13.5, d41.0, n13.9, k22.6, n12, d10.3, d56.1, k22.2, d10.0, n20, k57.4, d21.3, t91.1, d21.6, f52.9, s93.4, g50.8, s82.8, n11.1, d12.9, m35.8, d58.0, d12.0, e10.8, n63, m00.9, g52.7, e78.0, e78.2, m22.3, k90.0, d37.7, l52, d37.6, c92.8, d37.1, s24.1, d37.5, d37.4, s24.0, m21.8, m21.7, i95.9, d48.1, d48.2, e22.0, c15.2, n41.1, d59.8, e22.8, z03.2, j36, a63.0, d59.0, d13.3, m32.0, m89.0, e11.9, m43.0, j18.1, s32.00, l10.2, l97, i62.9, m54.1, k29.1, e24.0, i40.9, k82.3, s46.1, f22.0, t79.3, c81.9, g44.3, c81.4, k71.2, l43.9, l43.8, c50.6, g11.3, g11.4, g57.9, s66.0, l30.3, c72.5, c72.1, i42.7, d17.3, l30.9, i77.0, c83.0, k62.9, d50.9, k62.3, d24, d61.9, k51.5, d32, s52.50, k86.9, i44.0, k86.8, i44.3, c74.1, e28.3, d48, e84.8, n80.5, c77.4, d59, i45.0, q78.8, q43.0, j32.3, t19.1, k76.4, g03.1, s72.00, j32.8, i69.0, n03.1, m93.1, m93.2, k65.0, b00.9, l95.8, l95.9, s82.10, g80.8, s62, g80.9, s52.6, m16.2, g93.9, k43.5, i47.2, i47.9, c44.5, s82, s82.80, m80.5, j34.8, s72.2, d42.0, t06.8, q54.8, s82.70, k56.6, k56.4, i49.1, i49.2, i49.9, e88.9, g60.3, q85.0, m25.3, q52.4, s06.3, n45.9, c68.0, c22.0, n34.2, d35.0, d35.3, f42.2, e79.0, m23.4, d46.9, a04.7, t02.6, z09.0, s92.30, m54, n43.3, n43.2, i71.2, g40.4, d68.8, i71.8, k26, t90.0, d22.5, n32.3, n32.2";
//        String[] split = s.split(", ");
//        Arrays.sort(split);
//        System.out.println(Arrays.toString(split));
//    }

    private void copyToDirectory(Path to, Path from) throws IOException {
        if (Files.isDirectory(to) && Files.exists(to)) {
            FileUtils.deleteDirectory(to.toFile());
            Files.createDirectories(to);
        }
        FileUtils.copyFile(from.toFile(), to.toFile());
    }

    private void checkSpecificService(Class<? extends ReportService> clazz) {
        for (ReportService report : reports) {
            if (report.getClass().equals(clazz)) {
                try {
                    report.process();
                } catch (Exception e) {
                    log.error(e.getMessage(), e);
                }
            }
        }
    }

//    private void neverUse() {
//        ExecutorService executorService = Executors.newFixedThreadPool(reports.size());
//        for (int i = 0; i < reports.size(); i++) {
//            int finalI = i;
//            executorService.execute(() -> {
//                try {
//                    reports.get(finalI).process();
//                } catch (Exception e) {
//                    log.error(e.getMessage(), e);
//                }
//            });
//        }
//        for (ReportService report : reports) {
//            report.process();
//        }
//
//        ExecutorService executorService = Executors.newFixedThreadPool(3);
//        executorService.execute(() -> {
//            try {
//                checkSpecificService(DailyFormFirstSheetService.class);
//                checkSpecificService(DailyFormSecondSheetService.class);
//                reportsSaving.saveToDailyForm();
//            } catch (Exception e) {
//                log.error(e.getMessage(), e);
//            }
//        });
//        executorService.execute(() -> {
//            try {
//                checkSpecificService(FirstFormBazaService.class);
//                checkSpecificService(FirstFormPostService.class);
//                reportsSaving.saveToFirstForm();
//            } catch (Exception e) {
//                log.error(e.getMessage(), e);
//            }
//        });
//        executorService.execute(() -> {
//            try {
//                checkSpecificService(FourteenFormService.class);
//                reportsSaving.saveToFourteenForm();
//            } catch (Exception e) {
//                log.error(e.getMessage(), e);
//            }
//        });
//    }
}
