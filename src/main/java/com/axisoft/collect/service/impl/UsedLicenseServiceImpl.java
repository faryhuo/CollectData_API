package com.axisoft.collect.service.impl;

import com.axisoft.collect.entites.ExistLicense;
import com.axisoft.collect.service.UsedLicenseService;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class UsedLicenseServiceImpl implements UsedLicenseService {
    final String SHEET_NAME="Used";

    @Override
    public void insertUsedData(List<ExistLicense> existLicenseList, Workbook workbook) {
        Sheet sheet=workbook.getSheet(SHEET_NAME);
        int lastRowNum=sheet.getLastRowNum();
        double maxIndex=0;
        int existRowNum=0;
        for(int i=3;i<=lastRowNum;i++){
            Row row=sheet.getRow(i);
            double no=0;
            if(row.getCell(2)==null || StringUtils.isBlank(row.getCell(2).getStringCellValue())){
                existRowNum=row.getRowNum();
                break;
            }
            if(row.getCell(0)!=null){
                no=row.getCell(0).getNumericCellValue();
            }
            if(no>maxIndex){
                maxIndex=no;
            }

        }
        for(int i=0;i<existLicenseList.size();i++){
            ExistLicense existLicense=existLicenseList.get(i);
            int currentIndex=existRowNum+i;
            maxIndex++;
//            if(currentIndex>lastRowNum){
//                sheet.createRow(currentIndex);
//                for(int j=0;j<10;j++){
//                    sheet.getRow(currentIndex).createCell(j);
//                }
//            }
            Row row=sheet.getRow(currentIndex);
            if(row==null){
                row=sheet.createRow(currentIndex);
                for(int j=0;j<10;j++){
                    row.createCell(j);
                }
            }
            int currentLineIndex=currentIndex+1;
            row.getCell(0).setCellValue(maxIndex);
            String categoryFormula=String.format("IF(ISNA(INDEX('All License'!C:C,MATCH(C%d,'All License'!D:D,0))), \"\", INDEX('All License'!C:C,MATCH(C%d,'All License'!D:D,0)))",currentLineIndex,currentLineIndex);
            row.getCell(1).setCellFormula(categoryFormula);
            row.getCell(2).setCellValue(existLicense.getProductKey());
            row.getCell(3).setCellValue(existLicense.getNumber());
            row.getCell(4).setCellValue(existLicense.getUsername());
            String existFormulat=String.format(
                    "IF(ISERROR(HYPERLINK(\"#Existing!E\"&MATCH(TRIM(C%d)&TRIM(F%d)&(D%d),Existing!J:J,0),\"Y\")),\"\",HYPERLINK(\"#Existing!E\"&MATCH(TRIM(C%d)&TRIM(F%d)&(D%d),Existing!J:J,0),\"Y\"))",
                    currentLineIndex,currentLineIndex,currentLineIndex,currentLineIndex,currentLineIndex,currentLineIndex);
            row.getCell(5).setCellValue(existLicense.getMachineName());
            row.getCell(6).setCellFormula(existFormulat);
            String comparisionFormulat=String.format("C%d&F%d&D%d",currentLineIndex,currentLineIndex,currentLineIndex);
            row.getCell(9).setCellFormula(comparisionFormulat);

        }

    }
}
