package com.jin.commons.poi.model;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * @ClassName: CellStyleOptions
 * @Description: 样式
 * @author wujinglei
 * @date 2014年6月10日 下午3:15:07
 *
 */
public final class CellStyleOptions {

    private String titleFont = "微软雅黑";

    private String dataFont = "微软雅黑 Light";

    private short titleSize = 12;

    private short dataSize = 11;

    private short titleFontColor = IndexedColors.BLACK.getIndex();

    private short dataFontColor = IndexedColors.BLACK.getIndex();

    private BorderStyle[] titleBorder = new BorderStyle[]{BorderStyle.THIN,BorderStyle.THIN,BorderStyle.THIN,BorderStyle.THIN};

    private BorderStyle[] dataBorder = new BorderStyle[]{BorderStyle.THIN,BorderStyle.THIN,BorderStyle.THIN,BorderStyle.THIN};

    private short titleForegroundColor = IndexedColors.WHITE.getIndex();

    private short dataForegroundColor = IndexedColors.WHITE.getIndex();

    public String getTitleFont() {
        return titleFont;
    }

    public void setTitleFont(String titleFont) {
        this.titleFont = titleFont;
    }

    public String getDataFont() {
        return dataFont;
    }

    public void setDataFont(String dataFont) {
        this.dataFont = dataFont;
    }

    public short getTitleSize() {
        return titleSize;
    }

    public void setTitleSize(short titleSize) {
        this.titleSize = titleSize;
    }

    public short getDataSize() {
        return dataSize;
    }

    public void setDataSize(short dataSize) {
        this.dataSize = dataSize;
    }

    public BorderStyle[] getTitleBorder() {
        return titleBorder;
    }

    public void setTitleBorder(BorderStyle[] titleBorder) {
        this.titleBorder = titleBorder;
    }

    public BorderStyle[] getDataBorder() {
        return dataBorder;
    }

    public void setDataBorder(BorderStyle[] dataBorder) {
        this.dataBorder = dataBorder;
    }

    public short getTitleForegroundColor() {
        return titleForegroundColor;
    }

    public void setTitleForegroundColor(short titleForegroundColor) {
        this.titleForegroundColor = titleForegroundColor;
    }

    public short getDataForegroundColor() {
        return dataForegroundColor;
    }

    public void setDataForegroundColor(short dataForegroundColor) {
        this.dataForegroundColor = dataForegroundColor;
    }

    public short getTitleFontColor() {
        return titleFontColor;
    }

    public void setTitleFontColor(short titleFontColor) {
        this.titleFontColor = titleFontColor;
    }

    public short getDataFontColor() {
        return dataFontColor;
    }

    public void setDataFontColor(short dataFontColor) {
        this.dataFontColor = dataFontColor;
    }
}
