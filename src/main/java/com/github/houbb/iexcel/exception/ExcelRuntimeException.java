package com.github.houbb.iexcel.exception;

/**
 * EXCEL 运行时异常
 * @author binbin.hou
 * date 2018/11/1 10:57
 * @since 0.0.1
 */
public class ExcelRuntimeException extends RuntimeException {

    private static final long serialVersionUID = -3192101385447647614L;

    public ExcelRuntimeException() {
    }

    public ExcelRuntimeException(String message) {
        super(message);
    }

    public ExcelRuntimeException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelRuntimeException(Throwable cause) {
        super(cause);
    }

    public ExcelRuntimeException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
