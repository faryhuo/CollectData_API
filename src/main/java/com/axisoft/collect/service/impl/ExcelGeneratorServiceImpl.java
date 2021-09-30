package com.axisoft.collect.service.impl;

import com.axisoft.collect.entites.ComputerInfo;
import com.axisoft.collect.entites.ExistLicense;
import com.axisoft.collect.entites.LicenseInfo;
import com.axisoft.collect.service.AllLicenseService;
import com.axisoft.collect.service.ExcelGeneratorService;
import com.axisoft.collect.service.ExistLicenseService;
import com.axisoft.collect.service.UsedLicenseService;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class ExcelGeneratorServiceImpl implements ExcelGeneratorService {

    @Autowired
    ExistLicenseService existLicenseService;

    @Autowired
    UsedLicenseService usedLicenseService;

    @Override
    public void generateExcel(List<ComputerInfo> computerInfos, InputStream inputStream,OutputStream outputStream) throws IOException {
        Workbook workbook = null;
        try {
            ZipSecureFile.setMinInflateRatio(-1.0d);
            workbook = WorkbookFactory.create(inputStream);
            List<ExistLicense> existLicenseList=existLicenseService.insertExistData(computerInfos, workbook);
            usedLicenseService.insertUsedData(existLicenseList,workbook);
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

    public void generateComputerInfoExcel(List<ComputerInfo> computerInfos,OutputStream outputStream) throws IOException {
        Workbook workbook = null;
        try {
            ZipSecureFile.setMinInflateRatio(-1.0d);
            writerComputerInfoToExcel(computerInfos,outputStream);
        }finally {
            if(workbook!=null){
                workbook.close();
            }
        }
    }

    public Boolean writerComputerInfoToExcel(List<ComputerInfo> computerInfoList, OutputStream outputStream) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Computer Information");
        Row nRow = sheet.createRow(0);
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setFontName("Cambria");
        font.setFontHeightInPoints((short)11);
        font.setBold(false);
        style.setFont(font);
        Boolean hasError=false;
        CellStyle headerStyle = wb.createCellStyle();
        Font headerFont = wb.createFont();
        headerFont.setFontName("Cambria");
        headerFont.setFontHeightInPoints((short)13);
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        String[] titleList=new String[]{"No.","Machine Name","User","Product","Product key"};
        for(int i=0;i<titleList.length;i++){
            Cell nCell = nRow.createCell(i);
            nCell.setCellValue(titleList[i]);
            nCell.setCellStyle(headerStyle);
        }
        int no=1;
        for(int i=0;i<computerInfoList.size();i++){
            ComputerInfo computerInfo=computerInfoList.get(i);
            if(!computerInfo.getConvert()){
                hasError=true;
                continue;
            }
            Map<String,String> licenses=computerInfo.getSoftwareLicenses();
            if(licenses!=null) {
                for (String key: computerInfo.getSoftwareLicenses().keySet()) {
                    Row row = sheet.createRow(no);
                    Cell noCell = row.createCell(0);
                    noCell.setCellValue(no);
                    noCell.setCellStyle(style);
                    Cell machineNameCell = row.createCell(1);
                    machineNameCell.setCellValue(computerInfo.getComputerName());
                    machineNameCell.setCellStyle(style);
                    Cell userCell = row.createCell(2);
                    userCell.setCellValue(computerInfo.getWindowLogon());
                    userCell.setCellStyle(style);
                    Cell productCell = row.createCell(3);
                    productCell.setCellValue(key);
                    productCell.setCellStyle(style);
                    Cell productKeyCell = row.createCell(4);
                    String productKey=computerInfo.getSoftwareLicenses().get(key);
                    productKeyCell.setCellValue(productKey);
                    productKeyCell.setCellStyle(style);
                    no++;
                }
            }else{
                Row row = sheet.createRow(no);
                Cell noCell = row.createCell(0);
                noCell.setCellValue(no);
                noCell.setCellStyle(style);
                Cell machineNameCell = row.createCell(1);
                machineNameCell.setCellValue(computerInfo.getComputerName());
                machineNameCell.setCellStyle(style);
                Cell userCell = row.createCell(2);
                userCell.setCellValue(computerInfo.getWindowLogon());
                userCell.setCellStyle(style);
                Cell productCell = row.createCell(3);
                productCell.setCellValue("Fail to get the licenses.");
                CellStyle errorStyle = wb.createCellStyle();
                Font font2 = wb.createFont();
                font2.setFontName("Cambria");
                font2.setFontHeightInPoints((short)11);
                font2.setColor((short)2);
                font2.setBold(true);
                errorStyle.setFont(font2);
                productCell.setCellStyle(style);
                no++;
            }
        }
        for(int i=0;i<6;i++){
            sheet.autoSizeColumn(i);
        }
        if(hasError){
            createErrorSheet(wb,computerInfoList);
        }
        try {
            wb.write(outputStream);
        }finally {
            outputStream.close();
            wb.close();
        }
        return true;
    }

    public void createErrorSheet(XSSFWorkbook wb,List<ComputerInfo> computerInfoList){
        Sheet sheet = wb.createSheet("Error");
        Row nRow = sheet.createRow(0);
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setFontName("Cambria");
        font.setFontHeightInPoints((short)11);
        font.setBold(true);
        font.setColor((short)2);
        style.setFont(font);
        String[] titleList=new String[]{"No.","File Name","Message"};
        for(int i=0;i<titleList.length;i++){
            Cell nCell = nRow.createCell(i);
            nCell.setCellValue(titleList[i]);
            nCell.setCellStyle(style);
        }
        int i=0;
        for(int j=0;j<computerInfoList.size();j++) {
            ComputerInfo computerInfo = computerInfoList.get(j);
            if (!computerInfo.getConvert()) {
                i++;
                Row row = sheet.createRow(i);
                Cell noCell = row.createCell(0);
                noCell.setCellValue(i);
                noCell.setCellStyle(style);
                Cell fileCell = row.createCell(1);
                fileCell.setCellValue(computerInfo.getFileName());
                fileCell.setCellStyle(style);
                Cell messageCell = row.createCell(2);
                messageCell.setCellValue(computerInfo.getMessage());
                messageCell.setCellStyle(style);
            }
        }
    }


}
