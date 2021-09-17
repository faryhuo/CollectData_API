package com.axisoft.collect.service.impl;

import com.axisoft.collect.entites.ComputerInfo;
import com.axisoft.collect.entites.ExistLicense;
import com.axisoft.collect.entites.LicenseInfo;
import com.axisoft.collect.service.AllLicenseService;
import com.axisoft.collect.service.ExcelGeneratorService;
import com.axisoft.collect.service.ExistLicenseService;
import com.axisoft.collect.service.UsedLicenseService;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
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



}
