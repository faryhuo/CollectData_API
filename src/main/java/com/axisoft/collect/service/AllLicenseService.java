package com.axisoft.collect.service;

import com.axisoft.collect.entites.LicenseInfo;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.util.List;

public interface AllLicenseService {
    List<LicenseInfo> getAllLicenseKeyInfo(Workbook workbook) throws IOException;
}
