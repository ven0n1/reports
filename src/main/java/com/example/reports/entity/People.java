package com.example.reports.entity;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class People {

    private int all;
    private int emergency;
    private int ambulance;
    private int days;
    private int died;
}
