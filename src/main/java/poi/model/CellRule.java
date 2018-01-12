package poi.model;

/**
 * @ClassName: CellRule
 * @Description: 是否是必填项
 * @author wujinglei
 * @date 2014年6月10日 下午5:41:34
 *
 */
public enum CellRule {
	REQUIRED("required"),
	EQUALSTO("equalsTo"),
	LONG("long"),
	INTEGER("integer"),
	DOUBLE("double"),
	DATEFORMAT("dateFormat");
	
	String value;  
	
	private CellRule( String value ) {  
		this.value = value;  
    }  
}
