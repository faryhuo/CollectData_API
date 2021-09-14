package com.axisoft.collect.service;

import com.axisoft.collect.entites.ExistLicense;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public interface UsedLicenseService {
    void insertUsedData(List<ExistLicense> existLicenseList, Workbook workbook);
}
