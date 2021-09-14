package com.axisoft.collect.service.impl;

import com.axisoft.collect.entites.ExistLicense;
import com.axisoft.collect.service.UsedLicenseService;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class UsedLicenseServiceImpl implements UsedLicenseService {
    @Override
    public void insertUsedData(List<ExistLicense> existLicenseList, Workbook workbook) {

    }
}
