package com.example.reports.entity;

import lombok.AllArgsConstructor;
import lombok.Builder;

@lombok.Data
@AllArgsConstructor
@Builder
public class Data {

    private People all;
    private int village;
    private People adult;
    private People old;
    private People child;

    public Data() {
        all = new People();
        adult = new People();
        old = new People();
        child = new People();
    }
}
