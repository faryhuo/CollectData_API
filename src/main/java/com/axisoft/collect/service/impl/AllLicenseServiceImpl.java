package com.axisoft.collect.service.impl;

import com.axisoft.collect.entites.LicenseInfo;
import com.axisoft.collect.service.AllLicenseService;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service
public class AllLicenseServiceImpl implements AllLicenseService {
    @Override
    public List<LicenseInfo> getAllLicenseKeyInfo(Workbook workbook) throws IOException {
        List<LicenseInfo> licenseInfoList=new ArrayList<>();
        try {
            Sheet sheet=workbook.getSheet("All License");
            int totalRow = sheet.getLastRowNum();
            for (int i = 2; i < totalRow; i++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    continue;
                }
                LicenseInfo licenseInfo =new LicenseInfo();
                String productName= row.getCell(2).getStringCellValue();
                String productKey= row.getCell(3).getStringCellValue();
                if(StringUtils.isNotBlank(productName) &&  StringUtils.isNotBlank(productKey) ) {
                    licenseInfo.setProductName(productName.trim());
                    licenseInfo.setProductKey(productKey.trim());
                    licenseInfoList.add(licenseInfo);
                }
            }
        }finally {
        }
        return licenseInfoList;
    }

}
