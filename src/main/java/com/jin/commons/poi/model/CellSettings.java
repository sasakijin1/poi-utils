package com.jin.commons.poi.model;

import org.apache.commons.lang3.ArrayUtils;

import java.util.List;
import java.util.Map;

/**
 * The type Cell settings.
 *
 * @author wujinglei
 * @ClassName: CellSettings
 * @Description: 单元格设置
 * @date 2014年6月10日 下午3:15:07
 */
public final class CellSettings {
	
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
	 * 公式设置
	 */
	private FormulaSettings formulaSettings;
	
	/**
	 * 固定值 
	 */
	private Boolean isFixedValue = false;

	/**
	 * 固定值内容
	 */
	private Map fixedMap;

	/**
	 * 是否是静态值
	 */
	private Boolean hasStaticValue = false;

	/**
	 * 静态值
	 */
	private String staticValue;

	/**
	 * 是否是下接选项
	 */
	private Boolean isSelect = false;

	/**
	 * 下拉选择配置
	 */
	private CellSelectSettings cellSelectSettings;

	/**
	 * 子CELL
	 */
	private CellSettings[] subCells;

	/**
	 * Cell对象类型，默认为自动
	 */
	private CellDataType cellDataType = CellDataType.AUTO;

	/**
	 * Cell样式
	 */
	private CellStyleSettings cellStyleSettings;

	private Class cellClass;

	private String[] formulaGroupNames;

	/**
	 *  日期格式
	 */
	private DatePattern pattern;

	private Boolean skip = false;

	/**
	 * cell顺序
	 */
	private Integer cellSeq;

	/**
	 * @author: wujinglei
	 * @date: 2014年6月10日 下午5:00:42
	 * @Description:隐藏构造方法
	 */
	private CellSettings(){
		
	}

	/**
	 * Instantiates a new Cell settings.
	 *
	 * @param key     :属性名
	 * @param colName :列名
	 * @author: wujinglei
	 * @date: 2014年6月10日 下午5:08:32
	 * @Description:设置关键字与显示列名
	 */
	public CellSettings(String key,String colName){
		this.key = key;
		this.colName = colName;
		this.cellStyleSettings = new CellStyleSettings();
	}

	/**
	 * Instantiates a new Cell settings.
	 *
	 * @param key           the key
	 * @param colName       the col name
	 * @param styleSettings the style settings
	 */
	public CellSettings(String key,String colName,CellStyleSettings styleSettings){
		this.key = key;
		this.colName = colName;
		this.cellStyleSettings = styleSettings;
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
	public Map getFixedMap() {
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
	 * Get sub cells cell settings [ ].
	 *
	 * @return the subCells
	 */
	public CellSettings[] getSubCells() {
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
	 * Add cell data type cell settings.
	 *
	 * @param cellDataType the cell data type
	 * @return cell settings
	 * @author: wujinglei
	 * @date: 2014 -8-13 下午5:49:06
	 * @Description: 设置读取XLS单元格数据类型 ，按设置类型读取，出错时记录异常，跳过添加数据
	 */
	public CellSettings addCellDataType(CellDataType cellDataType) {
		this.cellDataType = cellDataType;
		if (this.cellDataType == CellDataType.DATE){
			this.pattern = DatePattern.DATE_FORMAT_SEC;
		}
		return this;
	}

	/**
	 * Add cell rule cell settings.
	 *
	 * @param cellRule the cell rule
	 * @return the cell settings
	 */
	public CellSettings addCellRule(CellRule cellRule){
		this.cellRule = cellRule;
		return this;
	}

	/**
	 * Add cell rule cell settings.
	 *
	 * @param cellRule      the cell rule
	 * @param cellRuleValue the cell rule value
	 * @return cell settings
	 * @author: wujinglei
	 * @date: 2014 -8-13 下午5:50:43
	 * @Description: 添加CELL规则功能
	 */
	public CellSettings addCellRule(CellRule cellRule,Object cellRuleValue) {
		this.cellRule = cellRule;
		this.cellRuleValue = cellRuleValue;
		return this;
	}

	/**
	 * Add fixed map cell settings.
	 *
	 * @param fixedMap the fixed map
	 * @return cell settings
	 * @author: wujinglei
	 * @date: 2014 -8-13 下午5:51:17
	 * @Description: 添加固定项
	 */
	public CellSettings addFixedMap(Map fixedMap) {
		this.isFixedValue = true;
		this.fixedMap = fixedMap;
		return this;
	}

	/**
	 * Add static value cell settings.
	 *
	 * @param staticValue the static value
	 * @return cell settings
	 * @author: wujinglei
	 * @date: 2014 -8-13 下午5:51:31
	 * @Description: 添加静态项
	 */
	public CellSettings addStaticValue(String staticValue) {
		this.hasStaticValue = true;
		this.staticValue = staticValue;
		return this;
	}

	/**
	 * Add cell select cell settings.
	 *
	 * @param key        the key
	 * @param name       the name
	 * @param selectList the select list
	 * @return the cell settings
	 */
	public CellSettings addCellSelect(String key,String name,List selectList){
		if (selectList != null && !selectList.isEmpty()) {
			this.isSelect = true;
			this.cellSelectSettings = new CellSelectSettings(key, name, selectList);
		}
		return this;
	}

	/**
	 * Add cell select cell settings.
	 *
	 * @param array the array
	 * @return the cell settings
	 */
	public CellSettings addCellSelect(String[] array){
		this.isSelect = true;
		this.cellSelectSettings = new CellSelectSettings(array);
		return this;
	}

	/**
	 * Is select cell settings.
	 *
	 * @return the cell settings
	 */
	public CellSettings isSelect(){
		this.isSelect = true;
		if (this.cellSelectSettings == null){
			this.cellSelectSettings = new CellSelectSettings(new String[0]);
		}
		return this;
	}

	/**
	 * Set select bind cell settings.
	 *
	 * @param bindKey   the bind key
	 * @param targetKey the target key
	 * @return the cell settings
	 */
	public CellSettings setSelectBind(String bindKey,String targetKey){
		this.isSelect = true;
		if (this.cellSelectSettings == null){
			this.cellSelectSettings = new CellSelectSettings(new String[0]);
		}
		this.cellSelectSettings.setBind(bindKey,targetKey);
		return this;
	}

	/**
	 * Add cell select cell settings.
	 *
	 * @param map the map
	 * @return the cell settings
	 */
	public CellSettings addCellSelect(Map map){
		if (map != null && !map.isEmpty()){
			this.isSelect = true;
			this.cellSelectSettings = new CellSelectSettings(map);
		}
		return this;
	}

	/**
	 * Add sub cells cell settings.
	 *
	 * @param subCells the sub cells
	 * @return cell settings
	 * @author: wujinglei
	 * @date: 2014 -8-13 下午5:51:42
	 * @Description: 添加子项
	 */
	public CellSettings addSubCells(CellSettings[] subCells) {
		this.subCells = subCells;
		return this;
	}

	/**
	 * Add formula group name cell settings.
	 *
	 * @param formulaGroupName the formula group name
	 * @return the cell settings
	 */
	public CellSettings addFormulaGroupName(String formulaGroupName){
		if (this.formulaGroupNames == null){
			this.formulaGroupNames = new String[]{formulaGroupName};
		}else{
			ArrayUtils.add(this.formulaGroupNames,formulaGroupName);
		}
		return this;
	}

	/**
	 * Add formula group name cell settings.
	 *
	 * @param formulaGroupNames the formula group names
	 * @return the cell settings
	 */
	public CellSettings addFormulaGroupName(String[] formulaGroupNames){
		this.formulaGroupNames = formulaGroupNames;
		return this;
	}

	/**
	 * Add formula settings cell settings.
	 *
	 * @param formulaType      the formula type
	 * @param formulaGroupName the formula group name
	 * @return the cell settings
	 */
	public CellSettings addFormulaSettings(FormulaType formulaType,String formulaGroupName){
		this.cellDataType = CellDataType.FORMULA;
		this.formulaSettings = new FormulaSettings(formulaType);
		this.formulaGroupNames = new String[]{formulaGroupName};
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
	 * Gets cell style settings.
	 *
	 * @return the cell style settings
	 */
	public CellStyleSettings getCellStyleSettings() {
		return cellStyleSettings;
	}

	/**
	 * Get select text list string [ ].
	 *
	 * @return the string [ ]
	 */
	public String[] getSelectTextList(){
		return cellSelectSettings.getSelectText();
	}

	/**
	 * Get select value list string [ ].
	 *
	 * @return the string [ ]
	 */
	public String[] getSelectValueList(){
		return cellSelectSettings.getSelectValue();
	}

	/**
	 * Get select cascade flag boolean.
	 *
	 * @return the boolean
	 */
	public Boolean getSelectCascadeFlag(){
		if (this.cellSelectSettings == null){
			return false;
		}else {
			return cellSelectSettings.getCascadeFlag();
		}
	}

	/**
	 * Gets cell class.
	 *
	 * @return the cell class
	 */
	public Class getCellClass() {
		return cellClass;
	}

	/**
	 * Sets cell class.
	 *
	 * @param cellClass the cell class
	 */
	public void setCellClass(Class cellClass) {
		this.cellClass = cellClass;
	}

	/**
	 * Add pattern cell settings.
	 *
	 * @param pattern the pattern
	 * @return the cell settings
	 */
	public CellSettings addPattern(DatePattern pattern){
		this.pattern = pattern;
		return this;
	}

	/**
	 * Gets pattern.
	 *
	 * @return the pattern
	 */
	public DatePattern getPattern() {
		return pattern;
	}

	/**
	 * Get bing key string.
	 *
	 * @return the string
	 */
	public String getBingKey(){
		return this.cellSelectSettings.getBandKey();
	}

	/**
	 * Get select target key string.
	 *
	 * @return the string
	 */
	public String getSelectTargetKey(){
		return this.cellSelectSettings.getTargetKey();
	}

	/**
	 * Get select source list list.
	 *
	 * @return the list
	 */
	public List getSelectSourceList(){
		return this.cellSelectSettings.getSourceList();
	}

	/**
	 * Gets cell select settings.
	 *
	 * @return the cell select settings
	 */
	public CellSelectSettings getCellSelectSettings() {
		return cellSelectSettings;
	}

	/**
	 * Set cell data type.
	 *
	 * @param cellDataType the cell data type
	 */
	public void setCellDataType(CellDataType cellDataType){
		this.cellDataType = cellDataType;
	}

	/**
	 * Get formula group names string [ ].
	 *
	 * @return the string [ ]
	 */
	public String[] getFormulaGroupNames() {
		return formulaGroupNames;
	}

	/**
	 * Gets formula settings.
	 *
	 * @return the formula settings
	 */
	public FormulaSettings getFormulaSettings() {
		return formulaSettings;
	}

	/**
	 * Skip cell settings.
	 *
	 * @return the cell settings
	 */
	public CellSettings skip(){
		this.skip = true;
		return this;
	}

	/**
	 * Is skip boolean.
	 *
	 * @return the boolean
	 */
	public Boolean isSkip(){
		return this.skip;
	}

	/**
	 * Gets cell seq.
	 *
	 * @return the cell seq
	 */
	public Integer getCellSeq() {
		return cellSeq;
	}

	/**
	 * Sets cell seq.
	 *
	 * @param cellSeq the cell seq
	 */
	public void setCellSeq(Integer cellSeq) {
		this.cellSeq = cellSeq;
	}
}