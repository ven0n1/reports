package com.example.reports.service;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.List;

@Component
@Slf4j
@RequiredArgsConstructor
public class HelperForTemp {

    public void writeFile(Path path, List<String> strings) throws IOException {
        Files.deleteIfExists(path);
        Files.createFile(path);
        for (String string : strings) {
            Files.writeString(path, string, StandardOpenOption.APPEND);
        }
    }
}
