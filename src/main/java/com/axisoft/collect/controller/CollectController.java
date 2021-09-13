package com.axisoft.collect.controller;

import com.axisoft.collect.entites.ComputerInfo;
import com.axisoft.collect.entites.LicenseInfo;
import com.axisoft.collect.entites.ResponseEntity;
import com.axisoft.collect.service.CollectService;
import com.axisoft.collect.service.ExcelCheck;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.*;
import java.util.*;

@Controller
public class CollectController {

    @Autowired
    ExcelCheck excelCheck;

    @PostMapping("/downloadFile")
    public void downloadFile(HttpServletResponse response, HttpSession session) throws IOException {

        response.setContentType("application/x-xls;charset = UTF-8");
        response.setHeader("Content-Disposition","attachment;filename=computerInfo.xls;");
        File excelFile=(File)session.getAttribute("tempTile");
        FileInputStream fileInputStream=new FileInputStream(excelFile);
        int len=0;
        byte[] buffer=new byte[1024];
        while((len=fileInputStream.read(buffer))!= -1){
            response.getOutputStream().write(buffer,0,len);
        }
    }

    @PostMapping("/download")
    @ResponseBody
    public ResponseEntity download(@RequestParam(name="files") MultipartFile[] files,@RequestParam(name="excelFile") MultipartFile[] excelFile, HttpSession session) throws IOException {
        CollectService collectService= new CollectService();
        Map<String,InputStream> inputStreams=new HashMap<>();
        Map<String,InputStream> excelInputStreams=new HashMap<>();

        for(int i=0;i<files.length;i++){
            inputStreams.put(files[i].getOriginalFilename(),files[i].getInputStream());
        }
        for(int i=0;i<excelFile.length;i++){
            excelInputStreams.put(excelFile[i].getOriginalFilename(),excelFile[i].getInputStream());
        }
        File tempTile=File.createTempFile("computerInfo",".xls");
        session.setAttribute("tempTile",tempTile);
        OutputStream outputStream=new FileOutputStream(tempTile);
        List<LicenseInfo> licenseInfo=collectService.getLicenseKeyInfo(excelInputStreams);
        List<ComputerInfo> computerInfoList=collectService.getComputerInfoList(inputStreams);
        collectService.writerComputerInfoToExcel(computerInfoList,licenseInfo,outputStream);
        return ResponseEntity.createSuccess();
    }

    @PostMapping("/checkExcelFile")
    @ResponseBody
    public ResponseEntity checkExcelFile(@RequestParam(name="excelFile") MultipartFile excelFile, HttpServletResponse response) throws IOException {
        List<String> messageList=excelCheck.validateFile(excelFile.getInputStream());
        if(messageList.size()>0){
            return ResponseEntity.createErrorByErrorMessage(messageList);
        }
        return ResponseEntity.createSuccess();
    }


}
