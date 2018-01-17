package com.jin.commons.poi.model;

import java.util.List;
import java.util.Map;

/**
 * The type Cell options.
 *
 * @author wujinglei
 * @ClassName: CellOptions
 * @Description: 单元格设置
 * @date 2014年6月10日 下午3:15:07
 */
public final class CellOptions {
	
	/**
	 * 对象的属性名
	 */
	private String key;
	
	/**
	 * 列名
	 */
	private String colName;
	
	/**
	 * 是否是规则
	 */
	private CellRule cellRule;
	
	/**
	 * 规则内容
	 */
	private Object cellRuleValue;
	
	/**
	 * 固定值 
	 */
	private Boolean isFixedValue = false;

	/**
	 * 固定值内容
	 */
	private Map<String,Object> fixedMap;
	
	private Boolean hasStaticValue = false;

	private Boolean isSelect = false;

	private CellSelect cellSelect;

	private String staticValue;
	
	private CellOptions[] subCells;
	
	private CellDataType cellDataType = CellDataType.AUTO;

	private CellStyleOptions cellStyleOptions;

	private Class cellClass;

	private DatePattern pattern;

	/**
	 * @author: wujinglei
	 * @date: 2014年6月10日 下午5:00:42
	 * @Description:隐藏构造方法
	 */
	@SuppressWarnings("unused")
	private CellOptions(){
		
	}

	/**
	 * Instantiates a new Cell options.
	 *
	 * @param key     :属性名
	 * @param colName :列名
	 * @author: wujinglei
	 * @date: 2014年6月10日 下午5:08:32
	 * @Description:设置关键字与显示列名
	 */
	public CellOptions(String key,String colName){
		this.key = key;
		this.colName = colName;
		this.cellStyleOptions = new CellStyleOptions();
	}

	/**
	 * Instantiates a new Cell options.
	 *
	 * @param key          the key
	 * @param colName      the col name
	 * @param styleOptions the style options
	 */
	public CellOptions(String key,String colName,CellStyleOptions styleOptions){
		this.key = key;
		this.colName = colName;
		this.cellStyleOptions = styleOptions;
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
	 * Gets col name.
	 *
	 * @return the colName
	 */
	public String getColName() {
		return colName;
	}

	/**
	 * Gets fixed map.
	 *
	 * @return the fixedMap
	 */
	public Map<String, Object> getFixedMap() {
		return fixedMap;
	}

	/**
	 * Gets cell rule.
	 *
	 * @return the cellRule
	 */
	public CellRule getCellRule() {
		return cellRule;
	}

	@Override
	public String toString() {
		return 
				"key : " + this.key + ", " + 
				"colName : " + this.colName + ", " + 
				"isFixedValue : " + this.isFixedValue + ", " +
				"cellRule : " + this.cellRule.value;
	}

	/**
	 * Gets static value.
	 *
	 * @return the staticValue
	 */
	public String getStaticValue() {
		return staticValue;
	}

	/**
	 * Gets cell rule value.
	 *
	 * @return the cellRuleValue
	 */
	public Object getCellRuleValue() {
		return cellRuleValue;
	}

	/**
	 * Get sub cells cell options [ ].
	 *
	 * @return the subCells
	 */
	public CellOptions[] getSubCells() {
		return subCells;
	}

	/**
	 * Gets cell data type.
	 *
	 * @return the cellDataType
	 */
	public CellDataType getCellDataType() {
		return cellDataType;
	}

	/**
	 * Add cell data type cell options.
	 *
	 * @param cellDataType the cell data type
	 * @return cell options
	 * @author: wujinglei
	 * @date: 2014 -8-13 下午5:49:06
	 * @Description: 设置读取XLS单元格数据类型 ，按设置类型读取，出错时记录异常，跳过添加数据
	 */
	public CellOptions addCellDataType(CellDataType cellDataType) {
		this.cellDataType = cellDataType;
		return this;
	}

	/**
	 * Add cell rule cell options.
	 *
	 * @param cellRule      the cell rule
	 * @param cellRuleValue the cell rule value
	 * @return cell options
	 * @author: wujinglei
	 * @date: 2014 -8-13 下午5:50:43
	 * @Description: 添加CELL规则功能
	 */
	public CellOptions addCellRule(CellRule cellRule,Object cellRuleValue) {
		this.cellRule = cellRule;
		this.cellRuleValue = cellRuleValue;
		return this;
	}

	/**
	 * Add fixed map cell options.
	 *
	 * @param fixedMap the fixed map
	 * @return cell options
	 * @author: wujinglei
	 * @date: 2014 -8-13 下午5:51:17
	 * @Description: 添加固定项
	 */
	public CellOptions addFixedMap(Map<String, Object> fixedMap) {
		this.isFixedValue = true;
		this.fixedMap = fixedMap;
		return this;
	}

	/**
	 * Add static value cell options.
	 *
	 * @param staticValue the static value
	 * @return cell options
	 * @author: wujinglei
	 * @date: 2014 -8-13 下午5:51:31
	 * @Description: 添加静态项
	 */
	public CellOptions addStaticValue(String staticValue) {
		this.hasStaticValue = true;
		this.staticValue = staticValue;
		return this;
	}

	/**
	 * Add cell select cell options.
	 *
	 * @param key        the key
	 * @param name       the name
	 * @param selectList the select list
	 * @return the cell options
	 */
	public CellOptions addCellSelect(String key,String name,List selectList){
		if (selectList != null && !selectList.isEmpty()) {
			this.isSelect = true;
			this.cellSelect = new CellSelect(key, name, selectList);
			this.addCellDataType(CellDataType.SELECT);
		}
		return this;
	}

	public CellOptions isSelect(){
		this.isSelect = true;
		this.addCellDataType(CellDataType.SELECT);
		return this;
	}

	public CellOptions setSelectBind(String bindKey,String targetKey){
		this.isSelect = true;
		this.addCellDataType(CellDataType.SELECT);
		this.cellSelect.setBind(bindKey,targetKey);
		return this;
	}

	/**
	 * Add cell select cell options.
	 *
	 * @param map the map
	 * @return the cell options
	 */
	public CellOptions addCellSelect(Map map){
		if (map != null && !map.isEmpty()){
			this.isSelect = true;
			this.cellSelect = new CellSelect(map);
			this.addCellDataType(CellDataType.SELECT);
		}
		return this;
	}

	/**
	 * Add sub cells cell options.
	 *
	 * @param subCells the sub cells
	 * @return cell options
	 * @author: wujinglei
	 * @date: 2014 -8-13 下午5:51:42
	 * @Description: 添加子项
	 */
	public CellOptions addSubCells(CellOptions[] subCells) {
		this.subCells = subCells;
		return this;
	}

	/**
	 * Gets fixed value.
	 *
	 * @return the fixed value
	 */
	public Boolean getFixedValue() {
		return isFixedValue;
	}

	/**
	 * Gets has static value.
	 *
	 * @return the has static value
	 */
	public Boolean getHasStaticValue() {
		return hasStaticValue;
	}

	/**
	 * Gets select.
	 *
	 * @return the select
	 */
	public Boolean getSelect() {
		return isSelect;
	}

	/**
	 * Gets cell style options.
	 *
	 * @return the cell style options
	 */
	public CellStyleOptions getCellStyleOptions() {
		return cellStyleOptions;
	}

	/**
	 * Get select text list string [ ].
	 *
	 * @return the string [ ]
	 */
	public String[] getSelectTextList(){
		return cellSelect.getSelectText();
	}

	/**
	 * Get select value list string [ ].
	 *
	 * @return the string [ ]
	 */
	public String[] getSelectValueList(){
		return cellSelect.getSelectValue();
	}

	/**
	 * Get select cascade flag boolean.
	 *
	 * @return the boolean
	 */
	public Boolean getSelectCascadeFlag(){
		return cellSelect.getCascadeFlag();
	}

	public Class getCellClass() {
		return cellClass;
	}

	public void setCellClass(Class cellClass) {
		this.cellClass = cellClass;
	}

	public CellOptions addPattern(DatePattern pattern){
		this.pattern = pattern;
		return this;
	}

	public DatePattern getPattern() {
		return pattern;
	}

	public String getBingKey(){
		return this.cellSelect.getBandKey();
	}

	public String getSelectTargetKey(){
		return this.cellSelect.getTargetKey();
	}

	public List getSelectSourceList(){
		return this.cellSelect.getSourceList();
	}

	public void setCellDataType(CellDataType cellDataType){
		this.cellDataType = cellDataType;
	}
}