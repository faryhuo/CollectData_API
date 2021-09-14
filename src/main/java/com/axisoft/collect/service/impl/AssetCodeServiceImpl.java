package com.axisoft.collect.service.impl;

import com.axisoft.collect.entites.AssetCode;
import com.axisoft.collect.service.AssetCodeService;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;

import java.util.HashMap;
import java.util.Map;

@Service
public class AssetCodeServiceImpl implements AssetCodeService {
    final String SHEET_NAME="Asset Code List";

    private Map<String,AssetCode> assetCodeMap=new HashMap<>();

    void init(Workbook workbook){
        if(assetCodeMap.size()==0) {
            Sheet sheet = workbook.getSheet(SHEET_NAME);
            int totalRowNum = sheet.getLastRowNum();
            for (int i = 1; i <= totalRowNum; i++) {
                Row row = sheet.getRow(i);
                if (row != null && row.getCell(0) != null && StringUtils.isNotBlank(row.getCell(0).getStringCellValue())) {
                    String inventoryCode = row.getCell(0).getStringCellValue();
                    String username = row.getCell(3).getStringCellValue();
                    AssetCode assetCode = new AssetCode();
                    assetCode.setUsername(username);
                    assetCode.setInventoryCode(inventoryCode);
                    assetCodeMap.put(inventoryCode,assetCode);
                }
            }
        }
    }


    @Override
    public AssetCode getAssetCodeByMachineName(String machineName,Workbook workbook) {
        if(StringUtils.isBlank(machineName)){
            return null;
        }
        init(workbook);
        String inventoryCodeInput=machineName.replaceAll("[(]in [A-Za-z0-9_-]+[)]","").trim();
        if (assetCodeMap.containsKey(inventoryCodeInput)) {
            return assetCodeMap.get(inventoryCodeInput);
        }
        return null;
    }
}
