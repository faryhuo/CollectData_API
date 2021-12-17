package com.axisoft.collect.service.impl;

import com.axisoft.collect.entites.ExistLicense;
import com.axisoft.collect.entites.LicenseInfo;
import com.axisoft.collect.service.ExcelUtilsService;
import com.axisoft.collect.service.UsedLicenseService;
import org.apache.commons.collections4.ListUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

@Service
public class UsedLicenseServiceImpl implements UsedLicenseService {
    final String SHEET_NAME="Used";
    org.slf4j.Logger logger= LoggerFactory.getLogger(ExistLicenseServiceImpl.class);

    public List<LicenseInfo> getAllLicenseKey(Workbook workbook){
        Sheet sheet = workbook.getSheet(SHEET_NAME);
        int lastRowNum = sheet.getLastRowNum();
        List<LicenseInfo> licenseInfoList=new ArrayList<>();
        for (int i = 3; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row.getCell(2) == null || StringUtils.isBlank(row.getCell(2).getStringCellValue())) {
                continue;
            }
            String productKey=row.getCell(2).getStringCellValue();
            String productName="";
//            if(row.getCell(1)!=null && StringUtils.isBlank(row.getCell(1).getStringCellValue().trim())){
//                productName=row.getCell(1).getStringCellValue();
//            }
            LicenseInfo licenseInfo=new LicenseInfo();
            licenseInfo.setProductKey(productKey);
            licenseInfo.setSeq(i);
            licenseInfo.setProductName(productName);
            licenseInfoList.add(licenseInfo);
        }
        return licenseInfoList;
    }


    @Autowired
    ExcelUtilsService excelUtilsService;

    private boolean checkLicense(List<String> productList,String license){
        if(productList.contains(license)){
            return false;
        }
        Pattern pattern = Pattern.compile("[A-Za-z0-9]{4,5}-[A-Za-z0-9]{4,5}-[A-Za-z0-9]{4,5}-[A-Za-z0-9]{4,5}-[A-Za-z0-9]{4,5}");
        boolean result= pattern.matcher(license).matches();
        if(result){
            productList.add(license);
        }
        return result;

    }


    public String getUniqueKey(String ProductKey,String assetCode){
        return ProductKey+"##"+assetCode;
    }

    @Override
    public void insertUsedData(List<ExistLicense> existLicenseList, Workbook workbook) {
        try {
            Sheet sheet = workbook.getSheet(SHEET_NAME);
            int lastRowNum = sheet.getLastRowNum();
            List<String> productList=new ArrayList<>();
            List<String> pkList=new ArrayList<>();

            double maxIndex = 0;
            int existRowNum = 0;
            List<CellStyle> cellStyles=new ArrayList<>();
            for (int i = 3; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                double no = 0;
                if (row.getCell(2) == null || StringUtils.isBlank(row.getCell(2).getStringCellValue())) {
                    continue;
                }
                if(cellStyles.size()==0){
                    for (int j = 0; j < 10; j++) {
                        Cell nCell=row.getCell(j);
                        if(nCell!=null){
                            cellStyles.add(nCell.getCellStyle());
                        }else{
                            cellStyles.add(excelUtilsService.getCellStyle(workbook));
                        }
                    }
                }
                String productKey=row.getCell(2).getStringCellValue();
                productList.add(productKey);
                String assetCode="";
                if(row.getCell(5)!=null && row.getCell(5).getStringCellValue()!=null && StringUtils.isBlank(row.getCell(5).getStringCellValue())){
                    assetCode=row.getCell(5).getStringCellValue().trim();
                }
                pkList.add(getUniqueKey(productKey,assetCode));
                existRowNum = row.getRowNum();
                if (row.getCell(0) != null) {
                    no = row.getCell(0).getNumericCellValue();
                }
                if (no > maxIndex) {
                    maxIndex = no;
                }

            }
            List<ExistLicense> arr= existLicenseList.stream().filter((value)->{
               return checkLicense(productList,value.getProductKey()) && !pkList.contains(getUniqueKey(value.getProductKey(),value.getMachineName()));
            }).collect(Collectors.toList());
            CellStyle cellStyle = excelUtilsService.getCellStyle(workbook);
            for (int i = 0; i < arr.size(); i++) {
                ExistLicense existLicense = arr.get(i);
                int currentIndex = existRowNum + i+1;
                maxIndex++;
//            if(currentIndex>lastRowNum){
//                sheet.createRow(currentIndex);
//                for(int j=0;j<10;j++){
//                    sheet.getRow(currentIndex).createCell(j);
//                }
//            }
                Row row = sheet.getRow(currentIndex);
                if (row == null) {
                    row = sheet.createRow(currentIndex);
                }
                for (int j = 0; j < 10; j++) {
                    Cell nCell=row.createCell(j);
                    if(j<cellStyles.size()){
                        nCell.setCellStyle(cellStyles.get(j));
                    }else{
                        nCell.setCellStyle(cellStyle);
                    }

                }
                int currentLineIndex = currentIndex + 1;
                row.getCell(0).setCellValue(maxIndex);
                String categoryFormula = String.format("IF(ISNA(INDEX('All License'!C:C,MATCH(C%d,'All License'!D:D,0))), \"\", INDEX('All License'!C:C,MATCH(C%d,'All License'!D:D,0)))", currentLineIndex, currentLineIndex);
                row.getCell(1).setCellFormula(categoryFormula);
                row.getCell(2).setCellValue(existLicense.getProductKey());
                row.getCell(3).setCellValue(existLicense.getNumber());
                String userFormula = String.format("IF(ISNA(INDEX('Asset Code List'!D:D,MATCH(F%d,'Asset Code List'!A:A,0))), \"\", INDEX('Asset Code List'!D:D,MATCH(F%d,'Asset Code List'!A:A,0)))"
                        , currentLineIndex, currentLineIndex);
                row.getCell(4).setCellFormula(userFormula);
                row.getCell(5).setCellValue(existLicense.getMachineName());
                String existFormulat = String.format(
                        "IF(ISERROR(HYPERLINK(\"#Existing!E\"&MATCH(TRIM(C%d)&TRIM(F%d)&(D%d),Existing!J:J,0),\"Y\")),\"\",HYPERLINK(\"#Existing!E\"&MATCH(TRIM(C%d)&TRIM(F%d)&(D%d),Existing!J:J,0),\"Y\"))",
                        currentLineIndex, currentLineIndex, currentLineIndex, currentLineIndex, currentLineIndex, currentLineIndex);
                row.getCell(6).setCellFormula(existFormulat);
                String comparisionFormulat = String.format("C%d&F%d&D%d", currentLineIndex, currentLineIndex, currentLineIndex);
                row.getCell(9).setCellFormula(comparisionFormulat);
            }
        }catch (Exception e){
           logger.error(e.getMessage(),e);
           throw new RuntimeException("Can not insert the data to Used sheet.");
        }
    }
}
