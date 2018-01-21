package com.jin.commons.poi.model;

/**
 * 公式类型
 * @author wujinglei
 */
public enum FormulaType {
    /**
     * 总合
     */
    SUM("SUM"),
    /**
     * 平均值
     */
    AVERAGE("AVERAGE"),
    /**
     * 计数
     */
    COUNT("COUNT"),
    /**
     * 最大值
     */
    MAX("MAX"),
    /**
     * 最小值
     */
    MIN("MIN");

    String value;

    private FormulaType(String value) {
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
