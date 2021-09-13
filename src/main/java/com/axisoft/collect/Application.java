package com.axisoft.collect;



import com.axisoft.collect.entites.ComputerInfo;
import com.axisoft.collect.entites.LicenseInfo;
import com.axisoft.collect.service.CollectService;
import org.apache.coyote.http11.AbstractHttp11Protocol;
import org.slf4j.LoggerFactory;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ApplicationContext;
import org.springframework.context.annotation.ComponentScan;

import javax.xml.ws.Endpoint;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.logging.Logger;

@SpringBootApplication
@ComponentScan(basePackages={"com.axisoft.collect"})
public class Application {


    public static boolean IS_OPEN=false;


    //Application entry
	public  static void main(String args[]) throws IOException {
	    if(args.length==3){
            CollectService collectService= new CollectService();
            FileOutputStream fileOutputStream=new FileOutputStream(new File(args[2]));
            List<ComputerInfo> computerInfoList=collectService.getComputerInfoList(args[0]);
            List<LicenseInfo> licenseInfoList=collectService.getLicenseKeyInfo(args[1]);
            collectService.writerComputerInfoToExcel(computerInfoList,licenseInfoList,fileOutputStream);
        }else {
            IS_OPEN=true;
            ApplicationContext applicationContext = SpringApplication.run(Application.class, args);
        }
    }


}
