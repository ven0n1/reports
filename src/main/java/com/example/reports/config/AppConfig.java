package com.example.reports.config;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Configuration;

import java.util.List;
import java.util.Map;

@Configuration
@ConfigurationProperties(prefix = "app")
@Data
public class AppConfig {

    private From from;
    private To to;
    private Map<String, Integer> age;
    private String tempValue;

    @Data
    public static class From {

        private List<Integer> baza;
        private List<Integer> post;
    }

    @Data
    public static class To {

        private Map<String, List<Integer>> dailyForm;
        private Map<String, List<Integer>> firstForm;
        private List<Integer> fourteenForm;
    }
}
