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
	private Map<String,Object> fixedMap;

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
	 * @param key          the key
	 * @param colName      the col name
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
	public CellSettings addFixedMap(Map<String, Object> fixedMap) {
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

	public CellSettings addCellSelect(String[] array){
		this.isSelect = true;
		this.cellSelectSettings = new CellSelectSettings(array);
		return this;
	}

	public CellSettings isSelect(){
		this.isSelect = true;
		if (this.cellSelectSettings == null){
			this.cellSelectSettings = new CellSelectSettings(new String[0]);
		}
		return this;
	}

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

	public CellSettings addFormulaGroupName(String formulaGroupName){
		if (this.formulaGroupNames == null){
			this.formulaGroupNames = new String[]{formulaGroupName};
		}else{
			ArrayUtils.add(this.formulaGroupNames,formulaGroupName);
		}
		return this;
	}

	public CellSettings addFormulaGroupName(String[] formulaGroupNames){
		this.formulaGroupNames = formulaGroupNames;
		return this;
	}

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

	public Class getCellClass() {
		return cellClass;
	}

	public void setCellClass(Class cellClass) {
		this.cellClass = cellClass;
	}

	public CellSettings addPattern(DatePattern pattern){
		this.pattern = pattern;
		return this;
	}

	public DatePattern getPattern() {
		return pattern;
	}

	public String getBingKey(){
		return this.cellSelectSettings.getBandKey();
	}

	public String getSelectTargetKey(){
		return this.cellSelectSettings.getTargetKey();
	}

	public List getSelectSourceList(){
		return this.cellSelectSettings.getSourceList();
	}

	public CellSelectSettings getCellSelectSettings() {
		return cellSelectSettings;
	}

	public void setCellDataType(CellDataType cellDataType){
		this.cellDataType = cellDataType;
	}

	public String[] getFormulaGroupNames() {
		return formulaGroupNames;
	}

	public FormulaSettings getFormulaSettings() {
		return formulaSettings;
	}

	public CellSettings skip(){
		this.skip = true;
		return this;
	}

	public Boolean isSkip(){
		return this.skip;
	}
}