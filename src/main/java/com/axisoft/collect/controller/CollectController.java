package com.axisoft.collect.controller;

import com.axisoft.collect.entites.ComputerInfo;
import com.axisoft.collect.entites.ResponseEntity;
import com.axisoft.collect.service.ComputerCollectService;
import com.axisoft.collect.service.ExcelGeneratorService;
import com.axisoft.collect.service.ExcelUtilsService;
import com.axisoft.collect.service.impl.ComputerCollectServiceImpl;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.*;
import java.util.*;

@Controller
@RequestMapping("/api")
public class CollectController {


    @Autowired
    ExcelUtilsService excelUtilsService;


    @Autowired
    ExcelGeneratorService excelGeneratorService;

    @Autowired
    ComputerCollectService computerCollectService;

    @PostMapping("/downloadFile")
    public void downloadFile(HttpServletResponse response, HttpSession session) throws IOException {

        response.setContentType("application/x-xls;charset = UTF-8");
        String tempTileName=(String)session.getAttribute("tempTileName");
        response.setHeader("Content-Disposition",String.format("attachment;filename=%s;",tempTileName));
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
    public ResponseEntity download(String computerInfoListStr,@RequestParam(name="excelFile") MultipartFile excelFile, HttpSession session) throws IOException {
        InputStream excelInputStreams=excelFile.getInputStream();
        File tempTile=File.createTempFile("computerInfo",".xls");
        session.setAttribute("tempTile",tempTile);
        session.setAttribute("tempTileName",excelFile.getOriginalFilename());
        OutputStream outputStream=new FileOutputStream(tempTile);
        ObjectMapper objectMapper =new ObjectMapper();
        List<ComputerInfo> computerInfoList=objectMapper.readValue(computerInfoListStr,new TypeReference<List<ComputerInfo>>() {});
        excelGeneratorService.generateExcel(computerInfoList,excelInputStreams,outputStream);
        return ResponseEntity.createSuccess();
    }
//    @PostMapping("/download")
//    @ResponseBody
//    public ResponseEntity download(@RequestParam(name="files") MultipartFile[] files,@RequestParam(name="excelFile") MultipartFile excelFile, HttpSession session) throws IOException {
//        Map<String,InputStream> inputStreams=new HashMap<>();
//        InputStream excelInputStreams=excelFile.getInputStream();
//
//        for(int i=0;i<files.length;i++){
//            inputStreams.put(files[i].getOriginalFilename(),files[i].getInputStream());
//        }
//
//        File tempTile=File.createTempFile("computerInfo",".xls");
//        session.setAttribute("tempTile",tempTile);
//        session.setAttribute("tempTileName",excelFile.getOriginalFilename());
//        OutputStream outputStream=new FileOutputStream(tempTile);
//        List<ComputerInfo> computerInfoList= computerCollectService.getComputerInfoList(inputStreams);
//        excelGeneratorService.generateExcel(computerInfoList,excelInputStreams,outputStream);
//        return ResponseEntity.createSuccess();
//    }

    @PostMapping("/checkExcelFile")
    @ResponseBody
    public ResponseEntity checkExcelFile(@RequestParam(name="excelFile") MultipartFile excelFile, HttpServletResponse response) throws IOException {
        List<String> messageList= excelUtilsService.validateFile(excelFile.getInputStream());
        if(messageList.size()>0){
            return ResponseEntity.createErrorByErrorMessage(messageList);
        }
        return ResponseEntity.createSuccess();
    }


    @PostMapping("/checkLicenseFile")
    @ResponseBody
    public ResponseEntity checkLicenseFile(@RequestParam(name="files") MultipartFile[] files, HttpServletResponse response, HttpSession session) throws IOException {
        Map<String,InputStream> inputStreams=new HashMap<>();
        for(int i=0;i<files.length;i++){
            inputStreams.put(files[i].getOriginalFilename(),files[i].getInputStream());
        }
        List<ComputerInfo> computerInfoList= computerCollectService.getComputerInfoList(inputStreams);
        session.setAttribute("computerInfoList",computerInfoList);
        return ResponseEntity.createSuccess(computerInfoList);
    }

}
