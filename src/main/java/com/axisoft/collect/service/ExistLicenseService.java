package com.axisoft.collect.service;

import com.axisoft.collect.entites.ComputerInfo;
import com.axisoft.collect.entites.ExistLicense;
import com.axisoft.collect.entites.LicenseInfo;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.util.List;

public interface ExistLicenseService {
    //String getMachineName(String machineName);
    //String getUserName(String machineName);
    String getProductKey(String key,String productName,List<LicenseInfo> licenseInfoList);
    List<ExistLicense> insertExistData(List<ComputerInfo> computerInfos, Workbook workbook) throws IOException;
}
