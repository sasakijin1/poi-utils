package com.jin.commons.poi.exception;

/**
 * 获取CELL值异常
 * Created by wujinglei on 2015/8/20.
 */
public class TableSettingsCheckException extends Exception{

    public TableSettingsCheckException() {
        super();
    }
    public TableSettingsCheckException(String msg) {
        super(msg);
    }
    public TableSettingsCheckException(String msg, Throwable cause) {
        super(msg, cause);
    }
    public TableSettingsCheckException(Throwable cause) {
        super(cause);
    }
}
