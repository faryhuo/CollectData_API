package com.axisoft.collect.service.impl;

import com.axisoft.collect.entites.ComputerInfo;
import com.axisoft.collect.entites.LicenseInfo;
import com.axisoft.collect.service.ComputerCollectService;
import com.axisoft.collect.utils.HTMLUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.htmlparser.Node;
import org.htmlparser.Parser;
import org.htmlparser.util.NodeIterator;
import org.htmlparser.util.ParserException;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ComputerCollectServiceImpl implements ComputerCollectService {

    private  String getElementData(String textName,String content) throws ParserException {
        Parser parser = Parser.createParser(content, HTMLUtils.ENCODE);
        for (NodeIterator i = parser.elements (); i.hasMoreNodes(); ) {
            Node node = i.nextNode();
            String value=getNodeDataByText(textName,node);
            if(value!=null){
                return value;
            }
        }
        return null;
    }

    private  String getNodeDataByText(String textName,Node node) throws ParserException {
        if(node.getChildren()!=null && node.getChildren().size()>0){
            for (NodeIterator i = node.getChildren().elements(); i.hasMoreNodes(); ) {
                Node childNode = i.nextNode();
                String value= getNodeDataByText(textName,childNode);
                if(value!=null){
                    return value;
                }
            }
        }else{
            if(textName.equals(node.getText())){
                Node nextNode= node.getParent().getNextSibling();
                if(nextNode!=null && nextNode.getLastChild()!=null){
                    return nextNode.getLastChild().getText();
                }
            }
        }
        return null;
    }

    public List<ComputerInfo> getComputerInfoList(String dir){
        List<ComputerInfo> computerInfoList=new ArrayList<ComputerInfo>();
        File folder=new File(dir);
        if(!folder.isDirectory()){
            return null;
        }
        String[] files=folder.list(new FilenameFilter(){
            @Override
            public boolean accept(File dir, String name) {
                return name.endsWith(".html") || name.endsWith(".htm");
            }
        });
        for(int i=0;i<files.length;i++){
            computerInfoList.add(getComputerInfo(dir+File.separator+files[i]));
        }
        return computerInfoList;
    }


    public List<ComputerInfo> getComputerInfoList(Map<String,InputStream> inputStreams){
        List<ComputerInfo> computerInfoList=new ArrayList<ComputerInfo>();
        for(String fileName:inputStreams.keySet()){
            computerInfoList.add(getComputerInfo(inputStreams.get(fileName),fileName));
        }
        return computerInfoList;
    }

    public List<LicenseInfo> getLicenseKeyInfo(Map<String,InputStream> inputStreams) throws IOException {
        for(String fileName:inputStreams.keySet()){
            return getLicenseKeyInfo(inputStreams.get(fileName));
        }
        return null;
    }
    public List<LicenseInfo> getLicenseKeyInfo(String path) throws IOException {
        return getLicenseKeyInfo(new File(path));
    }

    public List<LicenseInfo> getLicenseKeyInfo(File file) throws IOException {
        InputStream inputStream=new FileInputStream(file);
        return getLicenseKeyInfo(inputStream);
    }
    public List<LicenseInfo> getLicenseKeyInfo(InputStream inputStream) throws IOException {
        List<LicenseInfo> licenseInfoList=new ArrayList<>();
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(inputStream);
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
                    licenseInfo.setProductName(productName);
                    licenseInfo.setProductKey(productKey);
                    licenseInfoList.add(licenseInfo);
                }
            }
        }finally {
            if(inputStream!=null){
                inputStream.close();
            }
            if(workbook!=null){
                workbook.close();
            }
        }
        return licenseInfoList;
    }

    public ComputerInfo getComputerInfo(String path) {
        return getComputerInfo(new File(path));
    }
    public ComputerInfo getComputerInfo(InputStream inputStream,String fileName){
        ComputerInfo computerInfo=new ComputerInfo();
        computerInfo.setFileName(fileName);
        try {
            String content=HTMLUtils.readFile(inputStream);
            String computerName=getElementData("Computer Name:",content);
            String windowsLogon=getElementData("Windows Logon:",content);
            Map<String,String>  licenses= getSoftwareLicenses(content);
            computerInfo.setConvert(true);
            if(StringUtils.isBlank(computerName)){
                computerInfo.setConvert(false);
                computerInfo.addMessage("Not found computer name");
            }
            if(StringUtils.isBlank(windowsLogon)){
                computerInfo.setConvert(false);
                computerInfo.addMessage("Not found windows logon");
            }
            if(licenses==null || licenses.size()==0){
                computerInfo.setConvert(false);
                computerInfo.addMessage("Not found software licenses");
            }
            computerInfo.setComputerName(computerName);
            computerInfo.setWindowLogon(windowsLogon);
            computerInfo.setSoftwareLicenses(licenses);
        } catch (ParserException e) {
            computerInfo.setConvert(false);
            computerInfo.addMessage("File content fail to convert to html."+e.getMessage());
            e.printStackTrace();
        } catch (IOException e) {
            computerInfo.setConvert(false);
            computerInfo.addMessage("Fail to read the html."+e.getMessage());
            e.printStackTrace();
        }
        return computerInfo;
    }


    public ComputerInfo getComputerInfo(File file){
        ComputerInfo computerInfo=new ComputerInfo();
        computerInfo.setFileName(file.getName());
        try {
            String content=HTMLUtils.readFile(file);
            String computerName=getElementData("Computer Name:",content);
            String windowsLogon=getElementData("Windows Logon:",content);
            Map<String,String>  licenses= getSoftwareLicenses(content);
            computerInfo.setConvert(true);
            if(StringUtils.isBlank(computerName)){
                computerInfo.setConvert(false);
                computerInfo.addMessage("Not found computer name");
            }
            if(StringUtils.isBlank(windowsLogon)){
                computerInfo.setConvert(false);
                computerInfo.addMessage("Not found windows logon");
            }
            if(licenses==null || licenses.size()==0){
                computerInfo.setConvert(false);
                computerInfo.addMessage("Not found licenses");
            }
            computerInfo.setComputerName(computerName);
            computerInfo.setWindowLogon(windowsLogon);
            computerInfo.setSoftwareLicenses(licenses);
        } catch (ParserException e) {
            computerInfo.setConvert(false);
            computerInfo.addMessage("Fail to read the html."+e.getMessage());
            e.printStackTrace();
        } catch (IOException e) {
            computerInfo.setConvert(false);
            computerInfo.addMessage("Fail to read the html."+e.getMessage());
            e.printStackTrace();
        }
        return computerInfo;
    }


    private Map<String,String> getSoftwareLicenses(String content) throws ParserException {
        Parser parser = Parser.createParser(content, HTMLUtils.ENCODE);
        Node tableNode=null;
        for (NodeIterator i = parser.elements (); i.hasMoreNodes(); ) {
            Node node = i.nextNode();
            tableNode=getLicenesTableNode(node);
            if(tableNode!=null){
                break;
            }
        }
        if(tableNode!=null){
            return getLicenseInfo4Node(tableNode);
        }
        return null;
    }

    private Map<String,String>  getLicenseInfo4Node(Node node) throws ParserException {
        Map<String,String> licenses=new HashMap<String,String>();
        if(node.getChildren()!=null && node.getChildren().size()>0) {
            for (NodeIterator i = node.getChildren().elements(); i.hasMoreNodes(); ) {
                Node childNode = i.nextNode();
                if(childNode.toHtml().startsWith("<tr>")){
                    if(childNode.getChildren()!=null && childNode.getChildren().size()>0) {
                        for (NodeIterator tr = childNode.getChildren().elements(); tr.hasMoreNodes(); ) {
                            Node td = tr.nextNode();
                            licenses.put(td.toPlainTextString(),td.getNextSibling().toPlainTextString());
                            break;
                        }
                    }
                }
            }
        }
        return licenses;
    }


    private Node getLicenesTableNode(Node node) throws ParserException {
        String html=node.toHtml();
        if(html.startsWith("<div class=\"blurb\" id=\"licenses\">")){
            Node nextNode=node.getNextSibling();
            while(nextNode!=null){
                if(nextNode.toHtml().startsWith("<div class=\"reportSection\">")){
                    if(nextNode.getChildren()!=null && nextNode.getChildren().size()>0) {
                        for (NodeIterator div = nextNode.getChildren().elements(); div.hasMoreNodes(); ) {
                            Node divNode = div.nextNode();
                            if(divNode.toHtml().startsWith("<div class=\"reportSectionBody\">")) {
                                if(divNode.getChildren()!=null && divNode.getChildren().size()>0) {
                                    for (NodeIterator i = divNode.getChildren().elements(); i.hasMoreNodes(); ) {
                                        Node divChild = i.nextNode();
                                        if(divChild.toHtml().startsWith("<table role=\"presentation\">")) {
                                            return divChild;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                nextNode=nextNode.getNextSibling();
            }
        }
        if(node.getChildren()!=null && node.getChildren().size()>0){
            for (NodeIterator i = node.getChildren().elements(); i.hasMoreNodes(); ) {
                Node childNode = i.nextNode();
                Node tableNode=getLicenesTableNode(childNode);
                if(tableNode!=null){
                    return tableNode;
                }
            }
        }
        return null;
    }


    private String encodeKey(String key,String productName,List<LicenseInfo> licenseInfoList){
        Pattern pattern = Pattern.compile("[(]Key:([ 0-9a-zA-Z\\-,\\\\/']+)[)]");
        Matcher matcher =pattern.matcher(key);
        String encodeKey=key;
        while(matcher.find()) {
            if(matcher.group().length()>1) {
                encodeKey = matcher.group(1);
            }
        }
//        Map<String,String> map=new HashMap<>();
//        for(int i=0;i<licenseInfoList.size();i++){
//            LicenseInfo licenseInfo=licenseInfoList.get(i);
//            map.put(licenseInfo.)
//        }
        if(encodeKey.contains("ends with")){
            for(int i=0;i<licenseInfoList.size();i++){
                LicenseInfo licenseInfo=licenseInfoList.get(i);
                int startIndex=encodeKey.indexOf("ends with");
                int endIndex=encodeKey.indexOf(",");
                String endWithKey=null;
                if(endIndex!=-1 && endIndex>startIndex) {
                    endWithKey =encodeKey.substring(startIndex,endIndex);
                }else{
                    endWithKey =encodeKey.substring(startIndex);
                }
                endWithKey=endWithKey.replace("ends with","").trim();
                if(licenseInfo.getProductKey().endsWith(endWithKey) && productName.contains(licenseInfo.getProductName())){
                    return licenseInfo.getProductKey().trim();
                }
            }
            return encodeKey.trim();
        }else{
            return encodeKey.trim();
        }
    }


    public Boolean writerComputerInfoToExcel(List<ComputerInfo> computerInfoList,List<LicenseInfo> licenseInfoList, OutputStream outputStream) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Computer Information");
        Row nRow = sheet.createRow(0);
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setFontName("Cambria");
        font.setFontHeightInPoints((short)11);
        font.setBold(false);
        style.setFont(font);
        Boolean hasError=false;
        CellStyle headerStyle = wb.createCellStyle();
        Font headerFont = wb.createFont();
        headerFont.setFontName("Cambria");
        headerFont.setFontHeightInPoints((short)13);
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        String[] titleList=new String[]{"No.","Machine Name","User","Product","Product key","Product key 2"};
        for(int i=0;i<titleList.length;i++){
            Cell nCell = nRow.createCell(i);
            nCell.setCellValue(titleList[i]);
            nCell.setCellStyle(headerStyle);
        }
        int no=1;
        Map<String,Cell> productKeyMap=new HashMap<>();
        for(int i=0;i<computerInfoList.size();i++){
            ComputerInfo computerInfo=computerInfoList.get(i);
            if(!computerInfo.getConvert()){
                hasError=true;
                continue;
            }
            Map<String,String> licenses=computerInfo.getSoftwareLicenses();
            if(licenses!=null) {
                for (String key: computerInfo.getSoftwareLicenses().keySet()) {
                    Row row = sheet.createRow(no);
                    Cell noCell = row.createCell(0);
                    noCell.setCellValue(no);
                    noCell.setCellStyle(style);
                    Cell machineNameCell = row.createCell(1);
                    machineNameCell.setCellValue(computerInfo.getComputerName());
                    machineNameCell.setCellStyle(style);
                    Cell userCell = row.createCell(2);
                    userCell.setCellValue(computerInfo.getWindowLogon());
                    userCell.setCellStyle(style);
                    Cell productCell = row.createCell(3);
                    productCell.setCellValue(key);
                    productCell.setCellStyle(style);
                    Cell productKeyCell = row.createCell(4);
                    String productKey=computerInfo.getSoftwareLicenses().get(key);
                    productKeyCell.setCellValue(productKey);
                    productKeyCell.setCellStyle(style);
                    Cell productKey2Cell = row.createCell(5);
                    String encodeKey=encodeKey(productKey,key,licenseInfoList);
                    productKey2Cell.setCellValue(encodeKey);
                    if(productKeyMap.containsKey(encodeKey)){
                        CellStyle errorStyle = wb.createCellStyle();
                        Font font2 = wb.createFont();
                        font2.setFontName("Cambria");
                        font2.setFontHeightInPoints((short)11);
                        font2.setColor((short)2);
                        font2.setBold(true);
                        errorStyle.setFont(font2);
                        productKey2Cell.setCellStyle(errorStyle);
                        productKeyMap.get(encodeKey).setCellStyle(errorStyle);
                    }else {
                        productKey2Cell.setCellStyle(style);
                    }
                    productKeyMap.put(encodeKey,productKey2Cell);

                    no++;
                }
            }else{
                Row row = sheet.createRow(no);
                Cell noCell = row.createCell(0);
                noCell.setCellValue(no);
                noCell.setCellStyle(style);
                Cell machineNameCell = row.createCell(1);
                machineNameCell.setCellValue(computerInfo.getComputerName());
                machineNameCell.setCellStyle(style);
                Cell userCell = row.createCell(2);
                userCell.setCellValue(computerInfo.getWindowLogon());
                userCell.setCellStyle(style);
                Cell productCell = row.createCell(3);
                productCell.setCellValue("Fail to get the licenses.");
                CellStyle errorStyle = wb.createCellStyle();
                Font font2 = wb.createFont();
                font2.setFontName("Cambria");
                font2.setFontHeightInPoints((short)11);
                font2.setColor((short)2);
                font2.setBold(true);
                errorStyle.setFont(font2);
                productCell.setCellStyle(style);
                no++;
            }
        }
        for(int i=0;i<6;i++){
            sheet.autoSizeColumn(i);
        }
        if(hasError){
            createErrorSheet(wb,computerInfoList);
        }
        try {
            wb.write(outputStream);
        }finally {
            outputStream.close();
            wb.close();
        }
        return true;
    }

    public void createErrorSheet(XSSFWorkbook wb,List<ComputerInfo> computerInfoList){
        Sheet sheet = wb.createSheet("Error");
        Row nRow = sheet.createRow(0);
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setFontName("Cambria");
        font.setFontHeightInPoints((short)11);
        font.setBold(true);
        font.setColor((short)2);
        style.setFont(font);
        String[] titleList=new String[]{"No.","File Name","Message"};
        for(int i=0;i<titleList.length;i++){
            Cell nCell = nRow.createCell(i);
            nCell.setCellValue(titleList[i]);
            nCell.setCellStyle(style);
        }
        int i=0;
        for(int j=0;j<computerInfoList.size();j++) {
            ComputerInfo computerInfo = computerInfoList.get(j);
            if (!computerInfo.getConvert()) {
                i++;
                Row row = sheet.createRow(i);
                Cell noCell = row.createCell(0);
                noCell.setCellValue(i);
                noCell.setCellStyle(style);
                Cell fileCell = row.createCell(1);
                fileCell.setCellValue(computerInfo.getFileName());
                fileCell.setCellStyle(style);
                Cell messageCell = row.createCell(2);
                messageCell.setCellValue(computerInfo.getMessage());
                messageCell.setCellStyle(style);
            }
        }
    }


}
