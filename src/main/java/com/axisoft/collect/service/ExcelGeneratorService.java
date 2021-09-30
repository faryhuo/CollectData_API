package com.axisoft.collect.service;

import com.axisoft.collect.entites.ComputerInfo;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

public interface ExcelGeneratorService {
    void generateExcel(List<ComputerInfo> computerInfos, InputStream inputStream,OutputStream outputStream)  throws IOException;
    void generateComputerInfoExcel(List<ComputerInfo> computerInfos,OutputStream outputStream) throws IOException;
}
