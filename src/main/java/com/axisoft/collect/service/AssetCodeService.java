package com.axisoft.collect.service;

import com.axisoft.collect.entites.AssetCode;
import org.apache.poi.ss.usermodel.Workbook;

public interface AssetCodeService {
    AssetCode getAssetCodeByMachineName(String machineName,Workbook workbook);
}
