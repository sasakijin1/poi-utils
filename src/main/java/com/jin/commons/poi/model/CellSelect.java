package com.jin.commons.poi.model;

import com.jin.commons.poi.utils.BeanUtils;

import java.util.*;

/**
 * The type Cell select.
 *
 * @author wujinglei
 */
public class CellSelect {

    private String[] selectText;

    private String[] selectValue;

    private String mappingCode;

    private Boolean cascadeFlag;

    public CellSelect(String mappingCode) {
        this.mappingCode = mappingCode;
        this.cascadeFlag = true;
    }

    /**
     * Instantiates a new Cell select.
     *
     * @param text        the text
     * @param value       the value
     * @param selectList  the select list
     * @param mappingCode the mapping code
     */
    public CellSelect(String text, String value, List selectList,String mappingCode) {
        this.setKeyAndValue(text,value,selectList);
        this.mappingCode = mappingCode;
        this.cascadeFlag = true;
    }

    /**
     * Instantiates a new Cell select.
     *
     * @param text       the text
     * @param value      the value
     * @param selectList the select list
     */
    public CellSelect(String text, String value, List selectList) {
        this.setKeyAndValue(text,value,selectList);
        this.cascadeFlag = false;
    }

    /**
     * Instantiates a new Cell select.
     *
     * @param map         the map
     * @param mappingCode the mapping code
     */
    public CellSelect(Map<String, Object> map,String mappingCode) {
        this.setKeyAndValue(map);
        this.mappingCode = mappingCode;
        this.cascadeFlag = true;
    }

    /**
     * Instantiates a new Cell select.
     *
     * @param map the map
     */
    public CellSelect(Map<String, Object> map) {
        this.setKeyAndValue(map);
        this.cascadeFlag = false;
    }

    /**
     * Instantiates a new Cell select.
     *
     * @param arrays      the arrays
     * @param mappingCode the mapping code
     */
    public CellSelect(String[] arrays,String mappingCode) {
        this.setKeyAndValue(arrays);
        this.mappingCode = mappingCode;
        this.cascadeFlag = true;
    }

    /**
     * Instantiates a new Cell select.
     *
     * @param arrays the arrays
     */
    public CellSelect(String[] arrays) {
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
     * @param text       the text
     * @param value      the value
     * @param selectList the select list
     */
    private void setKeyAndValue(String text, String value, List selectList) {
        if (selectList != null && selectList.size() > 0) {
            this.selectText = new String[selectList.size()];
            this.selectValue = new String[selectList.size()];
            for (int i = 0; i < selectList.size(); i++) {
                Object obj = selectList.get(i);
                if (obj instanceof Map) {
                    this.selectText[i] = (String) ((Map) obj).get(text);
                    this.selectValue[i] = (String) ((Map) obj).get(value);
                } else {
                    this.selectText[i] = (String) BeanUtils.invokeGetter(obj, text);
                    this.selectValue[i] = (String) BeanUtils.invokeGetter(obj, value);
                }
            }
        }
    }

    /**
     * Instantiates a new Cell select.
     *
     * @param map the map
     */
    private void setKeyAndValue(Map<String, Object> map) {
        if (map != null && map.size() > 0) {
            this.selectText = new String[map.size()];
            this.selectValue = new String[map.size()];
            Set<String> keys = map.keySet();
            int index = 0;
            for (String key : keys) {
                this.selectText[index] = (String) map.get(key);
                this.selectValue[index] = key;
                index++;
            }
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
     * Gets mapping code.
     *
     * @return the mapping code
     */
    public String getMappingCode() {
        return mappingCode;
    }

    /**
     * Gets cascade flag.
     *
     * @return the cascade flag
     */
    public Boolean getCascadeFlag() {
        return cascadeFlag;
    }
}
