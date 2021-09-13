package com.axisoft.collect.entites;

import org.apache.commons.collections4.ListUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.util.ArrayUtil;

import java.util.List;

public class ResponseEntity<T> {
    private String message;
    private int status;

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public int getStatus() {
        return status;
    }

    public void setStatus(int status) {
        this.status = status;
    }

    public T getData() {
        return data;
    }

    public void setData(T data) {
        this.data = data;
    }

    private T data;


    public static ResponseEntity createSuccess(){
        ResponseEntity responseEntity=new ResponseEntity();
        responseEntity.setStatus(0);
        return responseEntity;
    }

    public static ResponseEntity createErrorByErrorMessage(List<String> messageList){
        ResponseEntity responseEntity=new ResponseEntity();
        responseEntity.setStatus(1);
        responseEntity.setMessage(StringUtils.join(messageList.toArray(),"\r\n"));
        return responseEntity;
    }
}
