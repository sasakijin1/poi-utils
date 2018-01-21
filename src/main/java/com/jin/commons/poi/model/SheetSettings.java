package com.jin.commons.poi.model;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.*;

/**
 * The type Sheet settings.
 *
 * @author wujinglei
 * @ClassName: SheetSettings
 * @Description: 表配置
 * @date 2014年6月11日 上午10:36:28
 */
public final class SheetSettings {
	
	/**
	 * 表名
	 */
	private String sheetName;
	
	/**
	 * 跳过条数
	 */
	private Integer skipRows;
	
	/**
	 * 表序号
	 */
	private Integer sheetSeq;
	
	/**
	 * 列设置
	 */
	private CellSettings[] cellSettings;
	
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

	private Map<String,List<String>> selectMap = new HashMap();

	/**
	 * The Select target set.
	 */
	public Set<String> selectTargetSet = new HashSet<String>();

	/**
	 * @author: wujinglei
	 * @date: 2014年6月11日 上午10:40:27
	 * @Description:
	 */
	@SuppressWarnings("unused")
	private SheetSettings(){

	}

	/**
	 * Instantiates a new Sheet settings.
	 *
	 * @param sheetSeq the sheet seq
	 * @param skipRows the skip rows
	 * @author: wujinglei
	 * @date: 2014年6月11日 上午10:40:41
	 * @Description:强制序号及忽略行数
	 */
	public SheetSettings(Integer sheetSeq,Integer skipRows){
		this.sheetSeq = sheetSeq;
		this.skipRows = skipRows;
	}

	/**
	 * Instantiates a new Sheet settings.
	 *
	 * @param sheetName the sheet name
	 */
	public SheetSettings(String sheetName){
		this.sheetName = sheetName;
	}

	/**
	 * Instantiates a new Sheet settings.
	 *
	 * @param sheetName the sheet name
	 * @param clazz     the clazz
	 */
	public SheetSettings(String sheetName,Class clazz){
		this.sheetName = sheetName;
		this.dataClazzType = clazz;
	}

	/**
	 * Instantiates a new Sheet settings.
	 *
	 * @param sheetName the sheet name
	 * @param sheetSeq  the sheet seq
	 * @param skipRows  the skip rows
	 * 强制序号及忽略行数
	 */
	public SheetSettings(String sheetName,Integer sheetSeq,Integer skipRows){
		this.sheetName = sheetName;
		this.sheetSeq = sheetSeq;
		this.skipRows = skipRows;
	}

	/**
	 * Instantiates a new Sheet settings.
	 *
	 * @param sheetName the sheet name
	 * @param sheetSeq  the sheet seq
	 * @param skipRows  the skip rows
	 * @param dataClazzType     the dataClazzType
	 * @author: wujinglei
	 * @date: 2014 -6-21 下午4:53:51
	 * @Description:按表名来构造(导入时用)
	 */
	public SheetSettings(String sheetName,Integer sheetSeq,Integer skipRows,Class dataClazzType){
		this.sheetName = sheetName;
		this.sheetSeq = sheetSeq;
		this.skipRows = skipRows;
		this.dataClazzType = dataClazzType;
	}

	/**
	 * Instantiates a new Sheet settings.
	 *
	 * @param sheetName     the sheet name
	 * @param exportData    the export data
	 * @param dataClazzType the data clazz type
	 * @author: wujinglei
	 * @date: 2014年6月11日 下午4:02:50
	 * @Description:(导出时用)
	 */
	public SheetSettings(String sheetName,List exportData,Class dataClazzType){
		this.sheetName = sheetName;
		this.exportData = exportData;
		this.dataClazzType = dataClazzType;
	}

	/**
	 * Add title sheet settings.
	 *
	 * @param title      the title
	 * @param titleStyle the title style
	 * @return the sheet settings
	 */
	public SheetSettings addTitle(String title,CellStyleSettings titleStyle){
		this.title = title;
		this.titleStyle = titleStyle;
		return this;
	}

	/**
	 * Add title sheet settings.
	 *
	 * @param title the title
	 * @return the sheet settings
	 */
	public SheetSettings addTitle(String title){
		this.title = title;
		CellStyleSettings cellStyleSettings = new CellStyleSettings();
		cellStyleSettings.setTitleFont("宋体");
		cellStyleSettings.setTitleSize((short) 20);
		cellStyleSettings.setTitleFontColor(IndexedColors.BLUE.getIndex());
		cellStyleSettings.setAlignment(HorizontalAlignment.CENTER);
		cellStyleSettings.setVerticalAlignment(VerticalAlignment.CENTER);
		this.titleStyle = cellStyleSettings;
		return this;
	}

	/**
	 * Gets sheet name.
	 *
	 * @return the sheetName
	 */
	public String getSheetName() {
		return sheetName;
	}

	/**
	 * Sets sheet name.
	 *
	 * @param sheetName the sheetName to set
	 */
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	/**
	 * Gets sheet seq.
	 *
	 * @return the sheetSeq
	 */
	public Integer getSheetSeq() {
		return sheetSeq;
	}

	/**
	 * Sets sheet seq.
	 *
	 * @param sheetSeq the sheetSeq to set
	 */
	public void setSheetSeq(Integer sheetSeq) {
		this.sheetSeq = sheetSeq;
	}

	/**
	 * Get cell settings cell settings [ ].
	 *
	 * @return the cellSettings
	 */
	public CellSettings[] getCellSettings() {
		return cellSettings;
	}

	/**
	 * Sets cell settings.
	 *
	 * @param cellSettings the cellSettings to set
	 */
	public void setCellSettings(CellSettings[] cellSettings) {
		this.cellSettings = cellSettings;
	}

	/**
	 * Set cell settings.
	 *
	 * @param list the list
	 */
	public void setCellSettings(List<CellSettings> list){
		CellSettings[] cellSettings = new CellSettings[1];
		this.cellSettings = list.toArray(cellSettings);
	}

	/**
	 * Gets skip rows.
	 *
	 * @return the skipRows
	 */
	public Integer getSkipRows() {
		return skipRows;
	}

	/**
	 * Sets skip rows.
	 *
	 * @param skipRows the skipRows to set
	 */
	public void setSkipRows(Integer skipRows) {
		this.skipRows = skipRows;
	}

	/**
	 * Gets export data.
	 *
	 * @return the exportData
	 */
	public List getExportData() {
		return exportData;
	}

	/**
	 * Sets export data.
	 *
	 * @param exportData the exportData to set
	 */
	public void setExportData(List exportData) {
		this.exportData = exportData;
	}

	/**
	 * Gets data clazz type.
	 *
	 * @return the dataClazzType
	 */
	public Class getDataClazzType() {
		return dataClazzType;
	}

	/**
	 * Gets cell address map.
	 *
	 * @return the cell address map
	 */
	public Map<String, String> getCellAddressMap() {
		return cellAddressMap;
	}

	/**
	 * Sets cell address map.
	 *
	 * @param cellAddressMap the cell address map
	 */
	public void setCellAddressMap(Map<String, String> cellAddressMap) {
		this.cellAddressMap = cellAddressMap;
	}

	/**
	 * Gets select map.
	 *
	 * @return the select map
	 */
	public Map<String, List<String>> getSelectMap() {
		return selectMap;
	}

	/**
	 * Gets select target set.
	 *
	 * @return the select target set
	 */
	public Set<String> getSelectTargetSet() {
		return selectTargetSet;
	}

	/**
	 * Sets data clazz type.
	 *
	 * @param dataClazzType the dataClazzType to set
	 */
	public void setDataClazzType(Class dataClazzType) {
		this.dataClazzType = dataClazzType;
	}
	
	@Override
	public String toString() {
		return 
				"sheetSeq : " + this.sheetSeq + ", " + 
				"sheetName : " + this.sheetName + ", " + 
				"skipRows : " + this.skipRows + ", " + 
				"dataClazzType : " + this.dataClazzType;
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
}
