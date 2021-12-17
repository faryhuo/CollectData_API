package com.axisoft.collect.service.impl;

import com.axisoft.collect.entites.AssetCode;
import com.axisoft.collect.entites.ComputerInfo;
import com.axisoft.collect.entites.ExistLicense;
import com.axisoft.collect.entites.LicenseInfo;
import com.axisoft.collect.service.*;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class ExistLicenseServiceImpl implements ExistLicenseService {

    org.slf4j.Logger logger= LoggerFactory.getLogger(ExistLicenseServiceImpl.class);

    @Autowired
    ExistLicenseService existLicenseService;

    @Autowired
    AllLicenseService allLicenseService;

    @Autowired
    ExcelUtilsService excelUtilsService;
    @Autowired
    UsedLicenseService usedLicenseService;
    @Autowired
    AssetCodeService assetCodeService;

    final String SHEET_NAME="Existing";

    private double maxNumber=0;

//    @Override
//    public String getMachineName(String machineName) {
//        return null;
//    }
//
//    @Override
//    public String getUserName(String machineName) {
//        return null;
//    }


    @Override
    public String getProductKey(String key,String productName,List<LicenseInfo> licenseInfoList){
        Pattern pattern = Pattern.compile("[(]Key:([ 0-9a-zA-Z\\-,\\\\/']+)[)]");
        Matcher matcher =pattern.matcher(key);
        String encodeKey=key;
        while(matcher.find()) {
            if(matcher.group().length()>1) {
                encodeKey = matcher.group(1);
            }
        }
        encodeKey=encodeKey.trim();
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
                if(licenseInfo.getProductKey().endsWith(contianKey)
                        && (productName.contains(licenseInfo.getProductName()) || StringUtils.isBlank(licenseInfo.getProductName()))){
                    return licenseInfo.getProductKey().trim();
                }
            }
            return encodeKey;

        } else if(encodeKey.endsWith("-XXXXX")){
            String contianKey=encodeKey.replaceAll("-XXXXX","");
            for(int i=0;i<licenseInfoList.size();i++) {
                LicenseInfo licenseInfo=licenseInfoList.get(i);
                if(licenseInfo.getProductKey().startsWith(contianKey)
                        && (productName.contains(licenseInfo.getProductName()) || StringUtils.isBlank(licenseInfo.getProductName()))){
                    return licenseInfo.getProductKey().trim();
                }
            }
            return encodeKey;
        }else{
            return encodeKey.trim();
        }
    }

    Map<String, Integer> productKeyMap = new HashMap<>();

    @Override
    public List<ExistLicense>  insertExistData(List<ComputerInfo> computerInfos, Workbook workbook) throws IOException {
        try {
            Sheet existingSheet = workbook.getSheet(SHEET_NAME);
            maxNumber = deleteSameMachineRecord(computerInfos, existingSheet, workbook);
            List<LicenseInfo> licenseInfoList = allLicenseService.getAllLicenseKeyInfo(workbook);
            int index = (int)maxNumber;
            List<ExistLicense> existLicenseList = new ArrayList<>();
            licenseInfoList.addAll(usedLicenseService.getAllLicenseKey(workbook));
            for (int j = 0; j < computerInfos.size(); j++) {
                ComputerInfo computerInfo = computerInfos.get(j);
                if (computerInfo.getSoftwareLicenses() == null) {
                    continue;
                }
                for (String key : computerInfo.getSoftwareLicenses().keySet()) {
                    ExistLicense existLicense = insertRecord(workbook, existingSheet, licenseInfoList, index, computerInfo, productKeyMap, key);
                    existLicenseList.add(existLicense);
                    index++;
                }
                productKeyMap.clear();
            }
            return existLicenseList;
        }catch (Exception e){
            logger.error(e.getMessage(),e);
            throw new RuntimeException("Can not insert the data to Existing sheet.");
        }
    }



    private ExistLicense insertRecord(Workbook workbook,Sheet existingSheet, List<LicenseInfo> licenseInfoList, int index, ComputerInfo computerInfo, Map<String, Integer> productKeyMap, String key) {
        try {
            ExistLicense existLicense = new ExistLicense();
            Row row = existingSheet.createRow(index);
            CellStyle cellStyle = excelUtilsService.getCellStyle(workbook);
            for (int i = 0; i < 8; i++) {
                Cell nCell = row.createCell(i);
                //cellStyle.setAlignment(HorizontalAlignment.CENTER);
                nCell.setCellStyle(cellStyle);
            }
            String productKey = computerInfo.getSoftwareLicenses().get(key);
            String encodeKey = existLicenseService.getProductKey(productKey, key, licenseInfoList);
            AssetCode assetCode = assetCodeService.getAssetCodeByMachineName(computerInfo.getComputerName(), workbook);
            maxNumber++;
            row.getCell(0).setCellValue(maxNumber);
            if (assetCode != null) {
                row.getCell(1).setCellValue(assetCode.getInventoryCode());
                row.getCell(2).setCellValue(assetCode.getUsername());
                existLicense.setMachineName(assetCode.getInventoryCode());
                existLicense.setUsername(assetCode.getUsername());
            } else {
                row.getCell(1).setCellValue(computerInfo.getComputerName());
                row.getCell(2).setCellValue(computerInfo.getWindowLogon());
                existLicense.setMachineName(computerInfo.getComputerName());
                existLicense.setUsername(computerInfo.getWindowLogon());
            }
            existLicense.setProduct(key);
            existLicense.setProductKey(encodeKey);
            row.getCell(3).setCellValue(key);
            row.getCell(4).setCellValue(encodeKey);
            String pk = encodeKey;
            if (productKeyMap.containsKey(pk)) {
                int number = productKeyMap.get(pk);
                number++;
                productKeyMap.put(pk, number);
                row.getCell(5).setCellValue(number);
                existLicense.setNumber(number);
            } else {
                productKeyMap.put(pk, 1);
                row.getCell(5).setCellValue(1);
                existLicense.setNumber(1);
            }
            int rowIndex = row.getRowNum() + 1;
            String foumula = String.format("IF(ISERROR(HYPERLINK(\"#Used!C\"&MATCH((E%d&B%d&F%d),Used!J:J,0),\"Y\")),\"\",HYPERLINK(\"#Used!C\"&MATCH((E%d&B%d&F%d),Used!J:J,0),\"Y\"))"
                    , rowIndex, rowIndex, rowIndex, rowIndex, rowIndex, rowIndex);
            row.getCell(6).setCellFormula(foumula);
            //row.getCell(6).getCellStyle().setAlignment(HorizontalAlignment.CENTER);
            //row.getCell(7).setCellValue(productKey);
            String comparisionFormulat = String.format("E%d&B%d&F%d", rowIndex, rowIndex, rowIndex);
            if(row.getCell(9)==null){
                row.createCell(9);
            }
            row.getCell(9).setCellFormula(comparisionFormulat);
            return existLicense;
        }catch (RuntimeException e){
            throw e;
        }catch (Exception e){
            logger.error(e.getMessage(),e);
            throw new RuntimeException("Can not insert the data to Existing sheet.");
        }
    }

    private double deleteSameMachineRecord(List<ComputerInfo> computerInfos, Sheet existingSheet,Workbook workbook) {
        try {
            int totalRow = existingSheet.getLastRowNum();
            List<Integer> needToDeleteRows = new ArrayList<>();
            List<List<Object>> oldData=new ArrayList<>();
            int currentIndex=1;
            for (int i = 2; i <= totalRow; i++) {
                Row row = existingSheet.getRow(i);
                if (row == null) {
                    needToDeleteRows.add(i);
                    continue;
                }
                String machineName = null;
                String user = null;
                double no = 0;
                if (row.getCell(0) != null) {
                    no = row.getCell(0).getNumericCellValue();
                }
                if (row.getCell(1) != null) {
                    machineName = row.getCell(1).getStringCellValue();
                }
                if (row.getCell(2) != null) {
                    user = row.getCell(2).getStringCellValue();
                }
                boolean hasExistRecord = false;
                if (StringUtils.isBlank(machineName) && StringUtils.isBlank(user)) {
                    needToDeleteRows.add(i);
                    continue;
                }
                for (int j = 0; j < computerInfos.size(); j++) {
                    ComputerInfo computerInfo = computerInfos.get(j);
                    AssetCode assetCode = assetCodeService.getAssetCodeByMachineName(computerInfo.getComputerName(), workbook);
                    if (assetCode != null) {
                        if ((StringUtils.isBlank(machineName) && StringUtils.isBlank(user) ||
                                (StringUtils.equals(assetCode.getInventoryCode(), machineName)
                                        && StringUtils.equals(assetCode.getUsername(), user)))) {
                            hasExistRecord = true;
                            break;
                        }
                    } else {
                        if ((StringUtils.isBlank(machineName) && StringUtils.isBlank(user) ||
                                (StringUtils.equals(computerInfo.getComputerName(), machineName)
                                        && StringUtils.equals(computerInfo.getWindowLogon(), user)))) {
                            hasExistRecord = true;
                            break;
                        }
                    }
                }
                //row.setZeroHeight(true);
                if (hasExistRecord) {
                    //row.setZeroHeight(true);
                    needToDeleteRows.add(i);
                }else{
                    List<Object> list=new ArrayList<>();
                    list.add(currentIndex);
                    for(int j=1;j<=4;j++) {
                        if (row.getCell(j) != null) {
                            list.add(row.getCell(j).getStringCellValue());
                        }else{
                            list.add("");
                        }
                    }
                    if (row.getCell(5) != null) {
                        list.add(row.getCell(5).getNumericCellValue());
                    }else{
                        list.add("");
                    }
                    oldData.add(list);
                    currentIndex++;
                }
            }

            insertOldData(oldData,existingSheet,workbook);
            for(int i=oldData.size()+2;i<totalRow;i++){
                if(existingSheet.getRow(i)!=null) {
                    existingSheet.removeRow(existingSheet.getRow(i));
                }
            }

            return oldData.size()+2;
        }catch (Exception e){
            logger.error(e.getMessage(),e);
            throw new RuntimeException("Can not delete the same data in Existing sheet.");
        }
    }

    private void insertOldData(List<List<Object>> oldData,Sheet existingSheet,Workbook workbook){
        CellStyle cellStyle = excelUtilsService.getCellStyle(workbook);
        //CellStyle centerCellStyle = excelUtilsService.getCellStyle(workbook,HorizontalAlignment.CENTER);
        for(int i=0;i<oldData.size();i++){
            int startIndex=2;
            if(existingSheet.getRow(startIndex+i)==null){
                existingSheet.createRow(startIndex+i);
            }
            Row row=existingSheet.getRow(startIndex+i);
            for(int j=0;j<7;j++){
                if(row.getCell(j)==null){
                    row.createCell(j);
                }
                if(j==0) {
                    row.getCell(j).setCellValue((Integer)oldData.get(i).get(j));
                }
                if(j>0 && j<=4) {
                    row.getCell(j).setCellValue((String)oldData.get(i).get(j));
                }
                if(j==4){
                    String pk = (String)oldData.get(i).get(4);
                    if (productKeyMap.containsKey(pk)) {
                        int number = productKeyMap.get(pk);
                        number++;
                        productKeyMap.put(pk, number);
                    } else {
                        productKeyMap.put(pk, 1);
                    }
                }
                if(j==5){
                    row.getCell(j).setCellValue((Double) oldData.get(i).get(j));
                }
                if(j==6) {
                    int rowIndex = row.getRowNum() + 1;
                    String foumula = String.format("IF(ISERROR(HYPERLINK(\"#Used!C\"&MATCH((E%d&B%d&F%d),Used!J:J,0),\"Y\")),\"\",HYPERLINK(\"#Used!C\"&MATCH((E%d&B%d&F%d),Used!J:J,0),\"Y\"))"
                            , rowIndex, rowIndex, rowIndex, rowIndex, rowIndex, rowIndex);
                    row.getCell(j).setCellFormula(foumula);
                }
                row.getCell(j).setCellStyle(cellStyle);

            }
            int rowIndex = row.getRowNum() + 1;
            String comparisionFormulat = String.format("E%d&B%d&F%d", rowIndex, rowIndex, rowIndex);
            if(row.getCell(9)==null){
                row.createCell(9);
            }
            row.getCell(9).setCellFormula(comparisionFormulat);
        }
    }


}
