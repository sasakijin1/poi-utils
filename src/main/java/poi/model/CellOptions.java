package poi.model;

import java.util.List;
import java.util.Map;

/**
 * @ClassName: CellOptions
 * @Description: 单元格设置
 * @author wujinglei
 * @date 2014年6月10日 下午3:15:07
 *
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
	 * 规则异常时，是否跳过异常
	 */
	private boolean isKeepInput;
	
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
	
	/**
	 * @author: wujinglei
	 * @date: 2014年6月10日 下午5:00:42
	 * @Description:隐藏构造方法
	 */
	@SuppressWarnings("unused")
	private CellOptions(){
		
	}
	
	/**
	 * @author: wujinglei
	 * @date: 2014年6月10日 下午5:08:32
	 * @Description:设置关键字与显示列名
	 * @param key:属性名
	 * @param colName:列名
	 */
	public CellOptions(String key,String colName){
		this.key = key;
		this.colName = colName;
		this.cellStyleOptions = new CellStyleOptions();
	}

	public CellOptions(String key,String colName,CellStyleOptions styleOptions){
		this.key = key;
		this.colName = colName;
		this.cellStyleOptions = styleOptions;
	}

	/**
	 * @return the key
	 */
	public String getKey() {
		return key;
	}

	/**
	 * @return the colName
	 */
	public String getColName() {
		return colName;
	}

	/**
	 * @return the fixedMap
	 */
	public Map<String, Object> getFixedMap() {
		return fixedMap;
	}

	/**
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
	 * @return the staticValue
	 */
	public String getStaticValue() {
		return staticValue;
	}

	/**
	 * @return the cellRuleValue
	 */
	public Object getCellRuleValue() {
		return cellRuleValue;
	}

	/**
	 * @return the subCells
	 */
	public CellOptions[] getSubCells() {
		return subCells;
	}

	/**
	 * @return the isKeepInput
	 */
	public boolean isKeepInput() {
		return isKeepInput;
	}

	/**
	 * @return the cellDataType
	 */
	public CellDataType getCellDataType() {
		return cellDataType;
	}
	
	/**
	 * @author: wujinglei
	 * @date: 2014-8-13 下午5:49:06
	 * @Description: 设置读取XLS单元格数据类型，按设置类型读取，出错时记录异常，跳过添加数据
	 * @param cellDataType
	 * @return
	 */
	public CellOptions addCellDataType(CellDataType cellDataType) {
		this.cellDataType = cellDataType;
		this.isKeepInput = false;
		return this;
	}

	/**
	 * @author: wujinglei
	 * @date: 2014-8-13 下午5:50:25
	 * @Description: 设置读取XLS单元格数据类型，按设置类型读取，出错时记录异常，按设置是否跳过添加数据
	 * @param cellDataType
	 * @param isKeepInput
	 * @return
	 */
	public CellOptions addCellDataType(CellDataType cellDataType,boolean isKeepInput) {
		this.cellDataType = cellDataType;
		this.isKeepInput = isKeepInput;
		return this;
	}

	/**
	 * @author: wujinglei
	 * @date: 2014-8-13 下午5:50:43
	 * @Description: 添加CELL规则功能
	 * @param cellRule
	 * @param cellRuleValue
	 * @param isKeepInput
	 * @return
	 */
	public CellOptions addCellRule(CellRule cellRule,Object cellRuleValue,boolean isKeepInput) {
		this.cellRule = cellRule;
		this.cellRuleValue = cellRuleValue;
		this.isKeepInput = isKeepInput;
		return this;
	}

	/**
	 * @author: wujinglei
	 * @date: 2014-8-13 下午5:51:17
	 * @Description: 添加固定项
	 * @param fixedMap
	 * @return
	 */
	public CellOptions addFixedMap(Map<String, Object> fixedMap) {
		this.isFixedValue = true;
		this.fixedMap = fixedMap;
		return this;
	}

	/**
	 * @author: wujinglei
	 * @date: 2014-8-13 下午5:51:31
	 * @Description: 添加静态项
	 * @param staticValue
	 * @return
	 */
	public CellOptions addStaticValue(String staticValue) {
		this.hasStaticValue = true;
		this.staticValue = staticValue;
		return this;
	}

	public CellOptions addCellSelect(String key,String name,List selectList){
		this.isSelect = true;
		this.cellSelect = new CellSelect(key,name,selectList);
		this.addCellDataType(CellDataType.SELECT);
		return this;
	}

	public CellOptions addCellSelect(Map map){
		this.isSelect = true;
		this.cellSelect = new CellSelect(map);
		this.addCellDataType(CellDataType.SELECT);
		return this;
	}

	/**
	 * @author: wujinglei
	 * @date: 2014-8-13 下午5:51:42
	 * @Description: 添加子项
	 * @param subCells
	 * @return
	 */
	public CellOptions addSubCells(CellOptions[] subCells) {
		this.subCells = subCells;
		return this;
	}

	public Boolean getFixedValue() {
		return isFixedValue;
	}

	public Boolean getHasStaticValue() {
		return hasStaticValue;
	}

	public Boolean getSelect() {
		return isSelect;
	}

	public String[] getSelectArray() {
		return cellSelect.getSelectArray();
	}

	public String getCellSelectValue(String key){
		return cellSelect.getValue(key);
	}

	public String getCellSelectRealValue(String key){
		return cellSelect.getRealValue(key);
	}

	public CellStyleOptions getCellStyleOptions() {
		return cellStyleOptions;
	}
}