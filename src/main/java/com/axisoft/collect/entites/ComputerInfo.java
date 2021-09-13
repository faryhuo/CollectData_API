package com.axisoft.collect.entites;


import org.apache.commons.lang3.StringUtils;

import java.util.Map;

public class ComputerInfo {
    private String fileName;
    private Boolean isConvert;
    private String message;
    private String computerName;
    private String windowLogon;
    private Map<String,String> softwareLicenses;

    public void addMessage(String msg){
        if(StringUtils.isNotBlank(message)){
            message=message+"\r\n"+msg;
        }else{
            message=msg;
        }
    }

    public String getComputerName() {
        return computerName;
    }

    public void setComputerName(String computerName) {
        this.computerName = computerName;
    }

    public String getWindowLogon() {
        return windowLogon;
    }

    public void setWindowLogon(String windowLogon) {
        this.windowLogon = windowLogon;
    }

    public Map<String, String> getSoftwareLicenses() {
        return softwareLicenses;
    }

    public void setSoftwareLicenses(Map<String, String> softwareLicenses) {
        this.softwareLicenses = softwareLicenses;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public Boolean getConvert() {
        return isConvert;
    }

    public void setConvert(Boolean convert) {
        isConvert = convert;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }
}
