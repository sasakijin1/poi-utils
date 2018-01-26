package com.jin.commons.poi.model;

import java.util.*;

/**
 * The type Table settings.
 */
public final class TableSettings {

    /**
     * 跳过条数
     */
    private Integer skipRows;

    /**
     * 表序号
     */
    private Integer tableSeq;

    /**
     * 列设置
     */
    private List<CellSettings> cellSettingsList;

    /**
     * 导出的数据
     */
    private List exportData;

    /**
     * 数据class类型
     */
    private Class dataClazzType;

    /**
     * 标题
     */
    private String title;

    private Integer cellCount;

    /**
     * 标题样式
     */
    private CellStyleSettings titleStyle;

    private Map<String,String> cellAddressMap = new HashMap();


    /**
     * Instantiates a new Table settings.
     */
    public TableSettings(){

    }

    /**
     * Instantiates a new Sheet settings.
     *
     * @param clazz the clazz
     */
    public TableSettings(Class clazz){
        this.dataClazzType = clazz;
    }

    /**
     * Instantiates a new Table settings.
     *
     * @param exportData the export data
     * @param clazz      the clazz
     */
    public TableSettings(List exportData,Class clazz){
        this.exportData = exportData;
        this.dataClazzType = clazz;
    }

    /**
     * Get cell settings cell settings [ ].
     *
     * @return the cellSettings
     */
    public List<CellSettings> getCellSettingsList() {
        return cellSettingsList;
    }

    /**
     * Sets cell settings.
     *
     * @param array the array
     */
    public void setCellSettings(CellSettings[] array) {
        this.cellSettingsList = new ArrayList<>(Arrays.asList(array));
    }

    /**
     * Sets cell settings.
     *
     * @param cellSettingsList the cell settings list
     */
    public void setCellSettings(List<CellSettings> cellSettingsList) {
        this.cellSettingsList = cellSettingsList;
    }

    /**
     * Sets export data.
     *
     * @param exportData the export data
     */
    public void setExportData(List exportData) {
        this.exportData = exportData;
    }

    /**
     * Gets cell count.
     *
     * @return the cell count
     */
    public Integer getCellCount() {
        return cellCount;
    }

    /**
     * Sets cell count.
     *
     * @param cellCount the cell count
     */
    public void setCellCount(Integer cellCount) {
        this.cellCount = cellCount;
    }

    /**
     * Gets skip rows.
     *
     * @return the skip rows
     */
    public Integer getSkipRows() {
        return skipRows;
    }

    /**
     * Gets table seq.
     *
     * @return the table seq
     */
    public Integer getTableSeq() {
        return tableSeq;
    }

    /**
     * Gets export data.
     *
     * @return the export data
     */
    public List getExportData() {
        return exportData;
    }

    /**
     * Gets data clazz type.
     *
     * @return the data clazz type
     */
    public Class getDataClazzType() {
        return dataClazzType;
    }

    /**
     * Gets title.
     *
     * @return the title
     */
    public String getTitle() {
        return title;
    }

    /**
     * Gets title style.
     *
     * @return the title style
     */
    public CellStyleSettings getTitleStyle() {
        return titleStyle;
    }

    /**
     * Gets cell address map.
     *
     * @return the cell address map
     */
    public Map<String, String> getCellAddressMap() {
        return cellAddressMap;
    }
}
