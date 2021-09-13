package com.axisoft.collect.config;

import com.axisoft.collect.Application;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.CommandLineRunner;
import org.springframework.context.annotation.Configuration;

@Configuration
public class RunConfig implements CommandLineRunner {

    @Value("${openProject.isOpen}")
    private boolean isOpen;

    @Value("${openProject.openUrl}")
    private String openUrl;

    @Value("${openProject.cmd}")
    private String cmd;

    @Override
    public void run(String... args){
        if(isOpen && Application.IS_OPEN){
            String runCmd = cmd + " " + openUrl ;
            Runtime run = Runtime.getRuntime();
            try {
                run.exec(runCmd);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
}
