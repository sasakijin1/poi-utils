package com.jin.commons.poi.exception;

/**
 * 获取CELL值异常
 * Created by wujinglei on 2015/8/20.
 */
public class CellGetOrSetException extends Exception{

    public CellGetOrSetException() {
        super();
    }
    public CellGetOrSetException(String msg) {
        super(msg);
    }
    public CellGetOrSetException(String msg, Throwable cause) {
        super(msg, cause);
    }
    public CellGetOrSetException(Throwable cause) {
        super(cause);
    }
}
