package com.example.reports.entity;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class SecondFormPeople {

    private int below14;
    private int between15_19;
    private int between20_24;
    private int between25_29;
    private int between30_34;
    private int between35_39;
    private int between40_44;
    private int between45_49;
    private int between50_54;
    private int between55_59;
    private int between60_64;
    private int between65_69;
    private int between70_74;
    private int between75_79;
    private int between80_84;
    private int above85;

    public void incBelow14() {
        below14++;
    }

    public void incBetween15_19() {
        between15_19++;
    }

    public void incBetween20_24() {
        between20_24++;
    }

    public void incBetween25_29() {
        between25_29++;
    }

    public void incBetween30_34() {
        between30_34++;
    }

    public void incBetween35_39() {
        between35_39++;
    }

    public void incBetween40_44() {
        between40_44++;
    }

    public void incBetween45_49() {
        between45_49++;
    }

    public void incBetween50_54() {
        between50_54++;
    }

    public void incBetween55_59() {
        between55_59++;
    }

    public void incBetween60_64() {
        between60_64++;
    }

    public void incBetween65_69() {
        between65_69++;
    }

    public void incBetween70_74() {
        between70_74++;
    }

    public void incBetween75_79() {
        between75_79++;
    }

    public void incBetween80_84() {
        between80_84++;
    }

    public void incAbove85() {
        above85++;
    }
}
