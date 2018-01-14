package poi.utils;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Collection;
import java.util.Map;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * The type Field utils.
 */
public class FieldUtils {

    /**
     * The constant log.
     */
    protected transient final static Logger log = LoggerFactory.getLogger(FieldUtils.class);

    /**
     * 实体的getXXX方法
     *
     * @param name 成员变量名
     * @return string string
     */
    public static String get(String name) {
        //get+变量名的第一个字母大写
        return "get" + (name.charAt(0) + "").toUpperCase() + name.substring(1);
    }

    /**
     * 实体的setXXX方法
     *
     * @param name 成员变量名
     * @return string string
     */
    public static String set(String name) {
        //get+变量名的第一个字母大写
        return "set" + (name.charAt(0) + "").toUpperCase() + name.substring(1);
    }

    /**
     * 取得一个成员变量的值
     *
     * @param tablebean tablebean
     * @param fieldName the field name
     * @return field value
     */
    public static Object getFieldValue(Object tablebean, String fieldName) {
        if (tablebean == null || fieldName == null || fieldName.trim().equals("")) {
            return null;
        } else if (Map.class.isAssignableFrom(tablebean.getClass())) {
            return ((Map) tablebean).get(fieldName);
        }
        //如果是数组直接返回其索引元素
        if(tablebean.getClass().isArray()){
            return ((Object[])tablebean)[Integer.parseInt(fieldName)];
        //如果是集合直接返回其索引元素
        }else if(Collection.class.isAssignableFrom(tablebean.getClass())){
            return ((Collection)tablebean).toArray()[Integer.parseInt(fieldName)];
        } else {
            try {
                //取得get方法
                Method method = tablebean.getClass().getMethod(get(fieldName), (Class[]) null);
                //调用实体类的getXXX方法
                return method.invoke(tablebean, (Object[]) (Class[]) null);
            } catch (NoSuchMethodException noSuchMethodException) {
                log.error("没有这个方法：" + get(fieldName), noSuchMethodException);
            } catch (IllegalAccessException illegalAccessException) {
            } catch (InvocationTargetException invocationTargetException) {
            }
        }
        return null;
    }

    /**
     * 循环获取子对象
     *
     * @param tablebean the tablebean
     * @param fieldName the field name
     * @return super field value
     */
    public static Object getSuperFieldValue(Object tablebean, String fieldName) {
        Object value = tablebean;
        if (fieldName == null || "".equals(fieldName.trim())) {
            return tablebean;
        }
        String[] split = fieldName.split("[.]");
        for (int i = 0; i < split.length - 1; i++) {
            value = getFieldValue(value, split[i]);
        }
        return getFieldValue(value, split[split.length - 1]);
    }

    /**
     * 循环获取子对象
     *
     * @param tablebean  源对象
     * @param fieldName  字段名
     * @param defaultVal 如果为null返回的值
     * @return object object
     */
    public static Object getSuperFieldValue(Object tablebean,String fieldName,Object defaultVal){
        Object superFieldValue = getSuperFieldValue(tablebean, fieldName);
        return superFieldValue==null?defaultVal:superFieldValue;
    }

    /**
     * 存入一个实体的成员变量值
     *
     * @param tablebean the tablebean
     * @param fieldName the field name
     * @param value     the value
     * @return field value
     */
    public static Object setFieldValue(Object tablebean, String fieldName, Object value) {
        if (value == null) {
            return tablebean;
        }
        try {
            if(tablebean instanceof Map){
                ((Map)tablebean).put(fieldName,value);
            }else {
                log.debug("准备为对象进行setter方法注入");
                //取得set方法
                Method method = tablebean.getClass().getMethod(set(fieldName), value.getClass());
                //调用实体类的setXXX方法
                method.invoke(tablebean, value);
            }
        } catch (IllegalAccessException ex) {
            log.error("FieldUtil错误：参数不正确", ex);
        } catch (NoSuchMethodException ex) {
            log.debug("为对象进行setter方法注入失败改为直接注入");
            BeanUtils.setFieldValue(tablebean, fieldName, value);
        } catch (InvocationTargetException ex) {
            log.error("FieldUtil：存入一个实体的成员变量值失败！", ex);
        }
        return tablebean;
    }

}