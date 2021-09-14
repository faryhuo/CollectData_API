package test.com.axisoft.collect.service.impl;

import com.axisoft.collect.service.ExcelCheckService;
import com.axisoft.collect.service.ExcelGeneratorService;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@SpringBootTest
public class ExcelGeneratorServiceImplTest  {

    @Autowired
    ExcelGeneratorService excelGeneratorService;


    void testGenerateExcel01(){
        excelGeneratorService.generateExcel(computerInfoList,excelInputStreams,outputStream);
    }
}
