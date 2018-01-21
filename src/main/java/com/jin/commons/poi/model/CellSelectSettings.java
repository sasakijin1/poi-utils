package com.jin.commons.poi.model;

import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * The type Cell select.
 *
 * @author wujinglei
 * @date 2018 -01-12
 */
public class CellSelectSettings {

    /**
     * 下拉文字
     */
    private String[] selectText;

    /**
     * 下拉值
     */
    private String[] selectValue;

    /**
     * 源列表
     */
    private List sourceList;

    /**
     * 绑定的Key
     */
    private String bandKey;

    /**
     * 联动下拉目标KEY
     */
    private String targetKey;

    /**
     * 是否子级联动下拉 父节点为FALSE
     */
    private Boolean cascadeFlag = false;

    private String key;

    private String name;

    /**
     * Sets bind.
     *
     * @param bandKey   the band key
     * @param targetKey the target key
     */
    public void setBind(String bandKey,String targetKey) {
        this.bandKey = bandKey;
        this.targetKey = targetKey;
        this.cascadeFlag = true;
    }

    /**
     * Instantiates a new Cell select.
     *
     * @param text       the text
     * @param value      the value
     * @param selectList the select list
     */
    public CellSelectSettings(String text, String value, List selectList) {
        this.sourceList = selectList;
        this.setKeyAndValue(text,value,selectList);
        this.cascadeFlag = false;
    }

    /**
     * Instantiates a new Cell select.
     *
     * @param map the map
     */
    public CellSelectSettings(Map<String, Object> map) {
        this.setKeyAndValue(map);
        this.cascadeFlag = false;
    }

    /**
     * Instantiates a new Cell select.
     *
     * @param arrays the arrays
     */
    public CellSelectSettings(String[] arrays) {
        this.setKeyAndValue(arrays);
        this.cascadeFlag = false;
    }

    private void setKeyAndValue(String[] arrays) {
        this.selectText = arrays;
        this.selectValue = arrays;
    }
    /**
     * Instantiates a new Cell select.
     *
     * @param key       the text
     * @param name      the value
     * @param selectList the select list
     */
    private void setKeyAndValue(String key, String name, List selectList) {
        this.key = key;
        this.name = name;
        this.sourceList = selectList;
    }

    /**
     * Instantiates a new Cell select.
     *
     * @param map the map
     */
    private void setKeyAndValue(Map<String, Object> map) {
        Set<String> keys = map.keySet();
        selectText = new String[keys.size()];
        selectValue = new String[keys.size()];
        int index = 0;
        for(String key: keys){
            selectValue[index] = key;
            selectText[index] = String.valueOf(map.get(key));
            index++;
        }
    }

    /**
     * Get select text string [ ].
     *
     * @return the string [ ]
     */
    public String[] getSelectText() {
        return selectText;
    }

    /**
     * Get select value string [ ].
     *
     * @return the string [ ]
     */
    public String[] getSelectValue() {
        return selectValue;
    }

    /**
     * Gets band key.
     *
     * @return the band key
     */
    public String getBandKey() {
        return bandKey;
    }

    /**
     * Gets target key.
     *
     * @return the target key
     */
    public String getTargetKey() {
        return targetKey;
    }

    /**
     * Gets cascade flag.
     *
     * @return the cascade flag
     */
    public Boolean getCascadeFlag() {
        return cascadeFlag;
    }

    /**
     * Gets source list.
     *
     * @return the source list
     */
    public List getSourceList() {
        return sourceList;
    }

    /**
     * Gets key.
     *
     * @return the key
     */
    public String getKey() {
        return key;
    }

    /**
     * Gets name.
     *
     * @return the name
     */
    public String getName() {
        return name;
    }
}
