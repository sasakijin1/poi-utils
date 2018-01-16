package com.jin.commons.poi.exception;

/**
 * 获取CELL值异常
 * Created by wujinglei on 2015/8/20.
 */
public class GetCellValueRunTimeException extends RuntimeException{

    public GetCellValueRunTimeException() {
        super();
    }
    public GetCellValueRunTimeException(String msg) {
        super(msg+ ":获取CELL值异常");
    }
    public GetCellValueRunTimeException(String msg, Throwable cause) {
        super(msg+ ":获取CELL值异常", cause);
    }
    public GetCellValueRunTimeException(Throwable cause) {
        super(cause);
    }
}
