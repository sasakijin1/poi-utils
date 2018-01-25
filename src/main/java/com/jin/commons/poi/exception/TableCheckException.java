package com.jin.commons.poi.exception;

/**
 * 获取CELL值异常
 * Created by wujinglei on 2015/8/20.
 */
public class TableCheckException extends Exception{

    public TableCheckException() {
        super();
    }
    public TableCheckException(String msg) {
        super(msg+ ":table 序号异常");
    }
    public TableCheckException(String msg, Throwable cause) {
        super(msg+ ":table 序号异常", cause);
    }
    public TableCheckException(Throwable cause) {
        super(cause);
    }
}
