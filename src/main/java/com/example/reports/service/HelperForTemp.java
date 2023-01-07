package com.example.reports.service;

import com.example.reports.util.PathsConstants;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FileUtils;
import org.springframework.stereotype.Component;

import javax.annotation.PostConstruct;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.List;

@Component
@Slf4j
@RequiredArgsConstructor
public class HelperForTemp {

//    @PostConstruct
    void init() throws IOException {
        createTempFile(PathsConstants.FOR_DAILY);
        createTempFile(PathsConstants.FOR_FIRST);
        createTempFile(PathsConstants.FOR_FOURTEEN);
    }

    public void createTempFile(Path path) throws IOException {
        FileUtils.copyFileToDirectory(PathsConstants.FROM_TEMP.toFile(), path.toFile());
    }

    public void writeFile(Path path, List<String> strings) throws IOException {
        Files.deleteIfExists(path);
        Files.createFile(path);
        for (String string : strings) {
            Files.writeString(path, string, StandardOpenOption.APPEND);
        }
    }
}
