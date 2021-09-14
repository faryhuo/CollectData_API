package com.axisoft.collect.service.impl;

import com.axisoft.collect.service.ExcelUtilsService;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelUtilsServiceImpl implements ExcelUtilsService {
    org.slf4j.Logger logger= LoggerFactory.getLogger(ExcelUtilsServiceImpl.class);

    public List<String> validateFile(InputStream inputStream) throws IOException {
        Workbook workbook = null;
        List<String> messageList=new ArrayList<>();
        String[] needToCheckSheetNameList=new String[]{"All License","Existing","Used","Asset Code List"};
        try {
            workbook = WorkbookFactory.create(inputStream);
            for(int i=0;i<needToCheckSheetNameList.length;i++) {
                Sheet sheet = workbook.getSheet(needToCheckSheetNameList[i]);
                if (sheet == null) {
                    messageList.add(String.format("Can not found the '%s' sheet in excel file", needToCheckSheetNameList[i]));
                }
            }
        } catch (IOException e) {
            logger.error(e.getMessage(),e);
            throw e;
        } finally {
            if(inputStream!=null){
                inputStream.close();
            }
            if(workbook!=null){
                workbook.close();
            }
        }
        return messageList;
    }

    public CellStyle getCellStyle(Workbook workbook){
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("Calibri");
        font.setFontHeightInPoints((short)11);
        font.setBold(false);
        style.setFont(font);
        return style;
    }



    public void removeRow(Sheet sheet, int rowIndex) {
        int lastRowNum=sheet.getLastRowNum();
        if(rowIndex>=0&&rowIndex<lastRowNum){
            sheet.shiftRows(rowIndex+1,lastRowNum, -1);
        }
        if(rowIndex==lastRowNum){
/*            Row removingRow=sheet.getRow(rowIndex);
            if(removingRow!=null){
                sheet.removeRow(removingRow);
            }*/
            sheet.shiftRows(lastRowNum+1,lastRowNum+1, -1);
        }
    }
}
