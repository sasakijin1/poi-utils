package com.jin.commons.poi.exception;

/**
 * 获取CELL值异常
 * Created by wujinglei on 2015/8/20.
 */
public class CellRuleException extends Exception{

    public CellRuleException() {
        super();
    }
    public CellRuleException(String msg) {
        super(msg);
    }
    public CellRuleException(String msg, Throwable cause) {
        super(msg, cause);
    }
    public CellRuleException(Throwable cause) {
        super(cause);
    }
}
