package com.allinmoney.platform.excel;

/**
 * Created by chris on 16/4/28.
 */
public class ExcelException extends RuntimeException{

    private Integer code;
    private String message;


    public ExcelException() {
        super();
    }

    public ExcelException(String message) {
        super(message);
        this.message = message;
    }

    public ExcelException(Integer code, String message) {
        super(message);
        this.code = code;
        this.message = message;
    }

    public Integer getCode() {
        return code;
    }

    public void setCode(Integer code) {
        this.code = code;
    }

    @Override
    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }


}
