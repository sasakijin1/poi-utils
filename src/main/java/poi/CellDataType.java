/**
*开发单位：FESCO Adecco 
*版权：FESCO Adecco
*@author：wujinglei
*@since： JDK1.6
*@version：1.0
*@date：2014-8-13 下午5:00:35
*/ 

package poi;

/**
 * @ClassName: CellDataType
 * @Description: CELL数据类型
 * @author wujinglei
 * @date 2014-8-13 下午5:00:35
 *
 */
public enum CellDataType {

	AUTO("auto"),
	VARCHAR("varchar"),
	NUMBER("number"),
	DATE("date"),
	TIMESTAMP("timestamp"),
	SELECT("select"),
	FORMULA("formula");
	
	String value;  
	
	private CellDataType( String value ) {  
		this.value = value;  
    }  
}

