package com.jin.commons.poi.utils;

import com.jin.commons.poi.model.CellDataType;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * The type Field utils.
 *
 * @author wujinglei
 */
public class FieldUtils {

    /**
     * The constant log.
     */
    protected transient final static Logger log = LoggerFactory.getLogger(FieldUtils.class);


    private static Map<Class<?>,CellDataType> cellDataTypeClazzMapping = new HashMap();

    static {
        cellDataTypeClazzMapping.put(Integer.class, CellDataType.INTEGER);
        cellDataTypeClazzMapping.put(Long.class, CellDataType.BIGINT);
        cellDataTypeClazzMapping.put(String.class, CellDataType.VARCHAR);
        cellDataTypeClazzMapping.put(BigDecimal.class, CellDataType.NUMBER);
        cellDataTypeClazzMapping.put(Double.class, CellDataType.NUMBER);
        cellDataTypeClazzMapping.put(Float.class, CellDataType.NUMBER);
        cellDataTypeClazzMapping.put(Boolean.class, CellDataType.BOOLEAN);
        cellDataTypeClazzMapping.put(Date.class, CellDataType.DATE);
    }

    /**
     * Get cell data type cell data type.
     *
     * @param clazz the clazz
     * @return the cell data type
     */
    public static CellDataType getCellDataType(Class clazz){
        return cellDataTypeClazzMapping.get(clazz);
    }

    /**
     * Get declared field field.
     *
     * @param clazz     the clazz
     * @param fieldName the field name
     * @return the field
     */
    public static Field getDeclaredField(Class clazz, String fieldName){
        Field field = null ;

        for(; clazz != Object.class ; clazz = clazz.getSuperclass()) {
            try {
                field = clazz.getDeclaredField(fieldName) ;
                return field ;
            } catch (Exception e) {

            }
        }

        return null;
    }

    /**
     * Get declared field type class.
     *
     * @param clazz     the clazz
     * @param fieldName the field name
     * @return the class
     */
    public static Class getDeclaredFieldType(Class clazz, String fieldName){
        Field field = getDeclaredField(clazz,fieldName);
        if (field != null){
            return field.getType();
        }else{
            return null;
        }
    }
}