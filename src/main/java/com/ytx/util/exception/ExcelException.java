package com.ytx.util.exception;

/**
 * @author wangqi
 * @version 1.0
 * @className ExcelException
 * @description TODO
 * @date 2019/7/8 11:37
 */
public class ExcelException extends RuntimeException{

    public ExcelException() {
    }

    public ExcelException(String message) {
        super(message);
    }

    public ExcelException(String message, Throwable cause) {
        super(message, cause);
    }
}
