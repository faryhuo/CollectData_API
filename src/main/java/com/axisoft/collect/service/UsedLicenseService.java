package com.axisoft.collect.service;

import com.axisoft.collect.entites.ExistLicense;
import com.axisoft.collect.entites.LicenseInfo;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.util.List;

public interface UsedLicenseService {
    void insertUsedData(List<ExistLicense> existLicenseList, Workbook workbook);

    List<LicenseInfo> getAllLicenseKey(Workbook workbook) throws IOException;
}
