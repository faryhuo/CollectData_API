package test.com.axisoft.collect.service;
import org.junit.Test;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class TestComputerCollectServiceImpl {

//    @Test
//    public void testGetComputerInfo01(){
//        ComputerCollectServiceImpl collectService= new ComputerCollectServiceImpl();
//        ComputerInfo computerInfo=collectService.getComputerInfo("C:\\Users\\faryhuo\\Desktop\\(L-01P-00-0003).html");
//        System.out.println(computerInfo.getComputerName());
//        System.out.println(computerInfo.getWindowLogon());
//        System.out.println(computerInfo.getSoftwareLicenses());
//    }

//    @Test
//    public void testGetComputerInfoList01(){
//        ComputerCollectServiceImpl collectService= new ComputerCollectServiceImpl();
//        List<ComputerInfo> computerInfoList=collectService.getComputerInfoList("C:\\Temp\\html");
//        for(int i=0;i<computerInfoList.size();i++) {
//            ComputerInfo computerInfo=computerInfoList.get(i);
//            System.out.println(computerInfo.getComputerName());
//            System.out.println(computerInfo.getWindowLogon());
//            System.out.println(computerInfo.getSoftwareLicenses());
//        }
//    }
//
//
//    @Test
//    public void testWriterComputerInfoToExcel() throws IOException {
//        ComputerCollectServiceImpl collectService= new ComputerCollectServiceImpl();
//        FileOutputStream fileOutputStream=new FileOutputStream(new File("D:\\temp\\excel\\excel.xls"));
//        List<ComputerInfo> computerInfoList=collectService.getComputerInfoList("C:\\Temp\\html");
//        //collectService.writerComputerInfoToExcel(computerInfoList,fileOutputStream);
//    }


    @Test
    public void testEncodeKey(){
        Pattern pattern = Pattern.compile("[(]Key:([ 0-9a-zA-Z\\-]+)[)]");
        Matcher matcher =pattern.matcher("02482-123-0000007-0XXXX (Key: XXXXX-JQR3C-2JRHY-XYRKY-QWPVM)");
        while(matcher.find()) {
            System.out.println(matcher.group(1));
        }
    }
}
