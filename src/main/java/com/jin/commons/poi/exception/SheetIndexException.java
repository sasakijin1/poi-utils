package com.jin.commons.poi.exception;

/**
 * 获取CELL值异常
 * Created by wujinglei on 2015/8/20.
 */
public class SheetIndexException extends Exception{

    public SheetIndexException() {
        super();
    }
    public SheetIndexException(String msg) {
        super(msg+ ":Sheet 序号异常");
    }
    public SheetIndexException(String msg, Throwable cause) {
        super(msg+ ":Sheet 序号异常", cause);
    }
    public SheetIndexException(Throwable cause) {
        super(cause);
    }
}
