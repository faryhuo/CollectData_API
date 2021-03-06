package com.axisoft.collect.service.impl;

import com.axisoft.collect.entites.AssetCode;
import com.axisoft.collect.entites.ComputerInfo;
import com.axisoft.collect.entites.ExistLicense;
import com.axisoft.collect.entites.LicenseInfo;
import com.axisoft.collect.service.AllLicenseService;
import com.axisoft.collect.service.AssetCodeService;
import com.axisoft.collect.service.ExcelUtilsService;
import com.axisoft.collect.service.ExistLicenseService;
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

//@Service
public class ExistLicense2ServiceImpl implements ExistLicenseService {

    org.slf4j.Logger logger= LoggerFactory.getLogger(ExistLicense2ServiceImpl.class);

    @Autowired
    ExistLicenseService existLicenseService;

    @Autowired
    AllLicenseService allLicenseService;

    @Autowired
    ExcelUtilsService excelUtilsService;

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

    @Override
    public List<ExistLicense>  insertExistData(List<ComputerInfo> computerInfos, Workbook workbook) throws IOException {
        try {
            Sheet existingSheet = workbook.getSheet(SHEET_NAME);
            maxNumber = deleteSameMachineRecord(computerInfos, existingSheet, workbook);
            int currentTotalRow = existingSheet.getLastRowNum();
            List<LicenseInfo> licenseInfoList = allLicenseService.getAllLicenseKeyInfo(workbook);
            int index = currentTotalRow + 1;
            List<ExistLicense> existLicenseList = new ArrayList<>();
            for (int j = 0; j < computerInfos.size(); j++) {
                ComputerInfo computerInfo = computerInfos.get(j);
                if (computerInfo.getSoftwareLicenses() == null) {
                    continue;
                }
                Map<String, Integer> productKeyMap = new HashMap<>();
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
            String pk = String.format("%s##%s##%s", computerInfo.getComputerName(), computerInfo.getWindowLogon(), encodeKey);
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
            //row.getCell(7).setCellValue(productKey);
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
            double maxNumber = 0;
            for (int i = 2; i <= totalRow; i++) {
                Row row = existingSheet.getRow(i);
                if (row == null) {
                    //existingSheet.createRow(i);
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
                if (maxNumber < no) {
                    maxNumber = no;
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
                }
            }
            Map<Integer,Integer> deleteParams=new HashMap<>();
            int lastRowNo=0;
            int currentRowNo=0;
            List<Integer> deleteParamsIndex=new ArrayList<>();
            for (int rowIndex =0; rowIndex <needToDeleteRows.size(); rowIndex++) {
                int rowNo=needToDeleteRows.get(rowIndex);
                if(rowNo-lastRowNo==1){
                    int count=deleteParams.get(currentRowNo);
                    count++;
                    deleteParams.put(currentRowNo,count);
                }else{
                    deleteParams.put(rowNo,1);
                    currentRowNo=rowNo;
                    deleteParamsIndex.add(rowNo);
                }
                lastRowNo=rowNo;
            }


            for (int rowIndex = deleteParamsIndex.size() - 1; rowIndex >= 0; rowIndex--) {
//                if(existingSheet.getRow(needToDeleteRows.get(rowIndex))!=null) {
//                    existingSheet.removeRow(existingSheet.getRow(needToDeleteRows.get(rowIndex)));
//                }
                excelUtilsService.removeRow(existingSheet,deleteParamsIndex.get(rowIndex),deleteParams.get(deleteParamsIndex.get(rowIndex)));

            }
            return maxNumber;
        }catch (Exception e){
            logger.error(e.getMessage(),e);
            throw new RuntimeException("Can not delete the same data in Existing sheet.");
        }
    }


}
