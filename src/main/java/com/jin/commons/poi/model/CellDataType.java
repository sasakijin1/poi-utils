package com.jin.commons.poi.model;

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
	INTEGER("integer"),
	BIGINT("bigint"),
	BOOLEAN("boolean"),
	DATE("date"),
	FORMULA("formula");
	
	String value;  
	
	private CellDataType( String value ) {  
		this.value = value;  
    }  
}

