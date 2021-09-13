package test.com.axisoft.collect.service;
import com.axisoft.collect.entites.ComputerInfo;
import com.axisoft.collect.service.CollectService;
import com.axisoft.collect.utils.HTMLUtils;
import org.htmlparser.Node;
import org.htmlparser.Parser;
import org.htmlparser.util.NodeIterator;
import org.htmlparser.util.ParserException;
import org.junit.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class TestCollectService {

//    @Test
//    public void testGetComputerInfo01(){
//        CollectService collectService= new CollectService();
//        ComputerInfo computerInfo=collectService.getComputerInfo("C:\\Users\\faryhuo\\Desktop\\(L-01P-00-0003).html");
//        System.out.println(computerInfo.getComputerName());
//        System.out.println(computerInfo.getWindowLogon());
//        System.out.println(computerInfo.getSoftwareLicenses());
//    }

//    @Test
//    public void testGetComputerInfoList01(){
//        CollectService collectService= new CollectService();
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
//        CollectService collectService= new CollectService();
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
