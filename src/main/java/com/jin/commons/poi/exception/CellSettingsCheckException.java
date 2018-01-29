package com.jin.commons.poi.exception;

/**
 * 获取CELL值异常
 * Created by wujinglei on 2015/8/20.
 */
public class CellSettingsCheckException extends Exception{

    public CellSettingsCheckException() {
        super();
    }
    public CellSettingsCheckException(String msg) {
        super(msg);
    }
    public CellSettingsCheckException(String msg, Throwable cause) {
        super(msg, cause);
    }
    public CellSettingsCheckException(Throwable cause) {
        super(cause);
    }
}
