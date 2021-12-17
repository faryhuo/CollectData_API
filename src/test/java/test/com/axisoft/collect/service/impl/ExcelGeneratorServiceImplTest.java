package test.com.axisoft.collect.service.impl;

import com.axisoft.collect.entites.ComputerInfo;
import com.axisoft.collect.service.ComputerCollectService;
import com.axisoft.collect.service.ExcelGeneratorService;
import com.axisoft.collect.service.impl.ComputerCollectServiceImpl;
import com.axisoft.collect.service.impl.ExcelGeneratorServiceImpl;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.stereotype.Service;
import org.springframework.test.context.junit4.SpringRunner;

import javax.annotation.Resources;
import java.io.*;
import java.net.URL;
import java.util.*;

@SpringBootTest(classes={com.axisoft.collect.Application.class})
@RunWith(SpringRunner.class)
public class ExcelGeneratorServiceImplTest  {

    @Autowired
    ExcelGeneratorService excelGeneratorService;

    @Autowired
    ComputerCollectService computerCollectService;


    @Test
    public void testGenerateExcel01() throws IOException {
        String inputExcelPath="D:\\temp\\Company Software License Test Template.xlsx";
        String outputExcelPath="D:\\temp\\excel_"+new Date().getTime()+".xlsx";
        testGenerateExcel(inputExcelPath,outputExcelPath);
    }

    @Test
    public void testGenerateExcel02() throws IOException {
        String inputExcelPath="D:\\temp\\Company Software License Test Template.xlsx";
        String outputExcelPath="D:\\temp\\excel_"+new Date().getTime()+".xlsx";
        testGenerateExcel(inputExcelPath,outputExcelPath);
        for(int i=0;i<10;i++){
            inputExcelPath=outputExcelPath;
            outputExcelPath="D:\\temp\\excel_"+new Date().getTime()+".xlsx";
            testGenerateExcel(inputExcelPath,outputExcelPath);

        }

    }

    private void testGenerateExcel(String inputExcelPath,String outputExcelPath) throws IOException {
        InputStream excelInputStreams=new FileInputStream(new File(inputExcelPath));
        URL resource = ExcelGeneratorServiceImplTest.class.getClassLoader().getResource("html");
        String path = "C:\\Temp\\html\\Software License & Hardware\\ALAB BLAB CLAB DLAB";//resource.getPath();
        File[] files=new File(path).listFiles();
        Map<String,InputStream> inputStreams=new HashMap<>();
        for(int i=0;i<files.length;i++) {
            inputStreams.put(files[i].getName(), new FileInputStream(files[i]));
        }
        List<ComputerInfo> computerInfoList= computerCollectService.getComputerInfoList(inputStreams);
        OutputStream outputStream=new FileOutputStream(new File(outputExcelPath));
        excelGeneratorService.generateExcel(computerInfoList,excelInputStreams,outputStream);
        outputStream.flush();
        outputStream.close();
    }
}
