package com.jin.commons.poi.model;

public enum DatePattern {

    DATE_FORMAT_DAY("yyyy-MM-dd"),
    DATE_FORMAT_DAY_2("yyyy/MM/dd"),
    TIME_FORMAT_SEC("HH:mm:ss"),
    DATE_FORMAT_SEC("yyyy-MM-dd HH:mm:ss"),
    DATE_FORMAT_MSEC("yyyy-MM-dd HH:mm:ss.SSS"),
    DATE_FORMAT_MSEC_T("yyyy-MM-dd'T'HH:mm:ss.SSS"),
    DATE_FORMAT_MSEC_T_Z("yyyy-MM-dd'T'HH:mm:ss.SSS'Z'"),
    DATE_FORMAT_DAY_SIMPLE("y/M/d");

    String value;

    private DatePattern(String value) {
        this.value = value;
    }

    @Override
    public String toString() {
        return this.value;
    }

    public String getValue(){
        return this.value;
    }
}
