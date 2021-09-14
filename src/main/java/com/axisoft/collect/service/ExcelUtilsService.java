package com.axisoft.collect.service;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public interface ExcelUtilsService {
    List<String> validateFile(InputStream inputStream)  throws IOException;
    CellStyle getCellStyle(Workbook workbook);
    void removeRow(Sheet sheet, int rowIndex);
}
