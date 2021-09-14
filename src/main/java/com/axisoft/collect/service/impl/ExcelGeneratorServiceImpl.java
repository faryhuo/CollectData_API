package com.axisoft.collect.service.impl;

import com.axisoft.collect.entites.ComputerInfo;
import com.axisoft.collect.entites.LicenseInfo;
import com.axisoft.collect.service.ExcelGeneratorService;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class ExcelGeneratorServiceImpl implements ExcelGeneratorService {

    @Override
    public void generateExcel(List<ComputerInfo> computerInfos, InputStream inputStream,OutputStream outputStream) throws IOException {
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(inputStream);
            Sheet existingSheet=workbook.getSheet("Existing");
            int totalRow = existingSheet.getLastRowNum();
            List<Integer> needToDeleteRows=new ArrayList<>();
            for (int i = 2; i < totalRow; i++) {
                Row row = existingSheet.getRow(i);
                if (row == null) {
                    continue;
                }
                String machineName=null;
                String user=null;
                if(row.getCell(1)!=null) {
                    machineName = row.getCell(1).getStringCellValue();
                }
                if(row.getCell(2)!=null) {
                    user = row.getCell(2).getStringCellValue();
                }
                boolean hasExistRecord=false;
                for(int j=0;j<computerInfos.size();j++){
                    ComputerInfo computerInfo=computerInfos.get(j);
                    if((StringUtils.isBlank(machineName) && StringUtils.isBlank(user)  ||
                            (StringUtils.equals(computerInfo.getComputerName(),machineName)
                                    && StringUtils.equals(computerInfo.getWindowLogon(),user)))){
                        hasExistRecord=true;
                        break;
                    }
                }
                if(hasExistRecord) {
                    needToDeleteRows.add(i);
                }
            }
            for(int rowIndex=needToDeleteRows.size()-1;rowIndex>=0;rowIndex--){
                removeRow(existingSheet,needToDeleteRows.get(rowIndex));
            }
            int currentTotalRow = existingSheet.getLastRowNum();
            List<LicenseInfo> licenseInfoList=getAllLicenseKeyInfo(workbook);
            for(int j=0;j<computerInfos.size();j++){
                ComputerInfo computerInfo = computerInfos.get(j);
                if(computerInfo.getSoftwareLicenses()==null){
                    continue;
                }
                for (String key: computerInfo.getSoftwareLicenses().keySet()) {
                    CellStyle style = getCellStyle(workbook);
                    Row row = existingSheet.createRow(currentTotalRow + j);
                    for (int i = 0; i < 8; i++) {
                        Cell nCell = row.createCell(i);
                        nCell.setCellStyle(style);
                    }
                    String productKey=computerInfo.getSoftwareLicenses().get(key);
                    String encodeKey=getProductKey(productKey,key,licenseInfoList);

                    row.getCell(1).setCellValue(computerInfo.getComputerName());
                    row.getCell(2).setCellValue(computerInfo.getWindowLogon());
                    row.getCell(3).setCellValue(key);
                    row.getCell(4).setCellValue(encodeKey);
                    row.getCell(7).setCellValue(productKey);

                }
            }
            workbook.write(outputStream);
        }finally {
            if(inputStream!=null){
                inputStream.close();
            }
            if(workbook!=null){
                workbook.close();
            }
        }
    }

    private List<LicenseInfo> getAllLicenseKeyInfo(Workbook workbook) throws IOException {
        List<LicenseInfo> licenseInfoList=new ArrayList<>();
        try {
            Sheet sheet=workbook.getSheet("All License");
            int totalRow = sheet.getLastRowNum();
            for (int i = 2; i < totalRow; i++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    continue;
                }
                LicenseInfo licenseInfo =new LicenseInfo();
                String productName= row.getCell(2).getStringCellValue();
                String productKey= row.getCell(3).getStringCellValue();
                if(StringUtils.isNotBlank(productName) &&  StringUtils.isNotBlank(productKey) ) {
                    licenseInfo.setProductName(productName.trim());
                    licenseInfo.setProductKey(productKey.trim());
                    licenseInfoList.add(licenseInfo);
                }
            }
        }finally {
        }
        return licenseInfoList;
    }

    private String getProductKey(String key,String productName,List<LicenseInfo> licenseInfoList){
        Pattern pattern = Pattern.compile("[(]Key:([ 0-9a-zA-Z\\-,\\\\/']+)[)]");
        Matcher matcher =pattern.matcher(key);
        String encodeKey=key.trim();
        while(matcher.find()) {
            if(matcher.group().length()>1) {
                encodeKey = matcher.group(1);
            }
        }
//        Map<String,String> map=new HashMap<>();
//        for(int i=0;i<licenseInfoList.size();i++){
//            LicenseInfo licenseInfo=licenseInfoList.get(i);
//            map.put(licenseInfo.)
//        }
        if(encodeKey.contains("ends with")){
            for(int i=0;i<licenseInfoList.size();i++){
                LicenseInfo licenseInfo=licenseInfoList.get(i);
                int startIndex=encodeKey.indexOf("ends with");
                int endIndex=encodeKey.indexOf(",");
                String endWithKey=null;
                if(endIndex!=-1 && endIndex>startIndex) {
                    endWithKey =encodeKey.substring(startIndex,endIndex);
                }else{
                    endWithKey =encodeKey.substring(startIndex);
                }
                endWithKey=endWithKey.replace("ends with","").trim();
                if(licenseInfo.getProductKey().endsWith(endWithKey) && productName.contains(licenseInfo.getProductName())){
                    return licenseInfo.getProductKey().trim();
                }
            }
            return encodeKey.trim();
        }else if(encodeKey.startsWith("XXXXX-")){
            String contianKey=encodeKey.replaceAll("XXXXX-","");
            for(int i=0;i<licenseInfoList.size();i++) {
                LicenseInfo licenseInfo=licenseInfoList.get(i);
                if(licenseInfo.getProductKey().endsWith(contianKey) && productName.contains(licenseInfo.getProductName())){
                    return licenseInfo.getProductKey().trim();
                }
            }
            return encodeKey;

        } else if(encodeKey.endsWith("-XXXXX")){
            String contianKey=encodeKey.replaceAll("-XXXXX","");
            for(int i=0;i<licenseInfoList.size();i++) {
                LicenseInfo licenseInfo=licenseInfoList.get(i);
                if(licenseInfo.getProductKey().startsWith(contianKey) && productName.contains(licenseInfo.getProductName())){
                    return licenseInfo.getProductKey().trim();
                }
            }
            return encodeKey;
        }else{
            return encodeKey.trim();
        }
    }

    private CellStyle getCellStyle(Workbook workbook){
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("Cambria");
        font.setFontHeightInPoints((short)11);
        font.setBold(false);
        style.setFont(font);
        return style;
    }

    private void removeRow(Sheet sheet, int rowIndex) {
        int lastRowNum=sheet.getLastRowNum();
        if(rowIndex>=0&&rowIndex<lastRowNum){
            sheet.shiftRows(rowIndex+1,lastRowNum, -1);
        }
        if(rowIndex==lastRowNum){
/*            Row removingRow=sheet.getRow(rowIndex);
            if(removingRow!=null){
                sheet.removeRow(removingRow);
            }*/
            sheet.shiftRows(lastRowNum,lastRowNum, -1);
        }
    }
}
