/**
*开发单位：FESCO Adecco 
*版权：FESCO Adecco
*@author：wujinglei
*@since： JDK1.6
*@version：1.0
*@date：2014年6月10日 下午5:41:34
*/ 

package poi;

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
