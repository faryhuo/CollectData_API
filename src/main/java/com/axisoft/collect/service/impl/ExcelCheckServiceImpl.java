package com.axisoft.collect.service.impl;

import com.axisoft.collect.service.ExcelCheckService;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelCheckServiceImpl implements ExcelCheckService {
    org.slf4j.Logger logger= LoggerFactory.getLogger(ExcelCheckServiceImpl.class);

    public List<String> validateFile(InputStream inputStream) throws IOException {
        Workbook workbook = null;
        List<String> messageList=new ArrayList<>();
        String[] needToCheckSheetNameList=new String[]{"All License","Existing","Used"};
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
}
