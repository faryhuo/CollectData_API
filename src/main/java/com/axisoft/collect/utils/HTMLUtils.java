package com.axisoft.collect.utils;

import org.htmlparser.Node;
import org.htmlparser.Parser;
import org.htmlparser.util.NodeIterator;
import org.htmlparser.util.ParserException;

import java.io.*;

public class HTMLUtils {
    public static String ENCODE = "UTF-8";

    public static String readFile(String path) throws IOException {
        BufferedReader bis=null;
        try {
             bis = new BufferedReader(new InputStreamReader(new FileInputStream( new File(path)), ENCODE) );
             String szContent="";
             String szTemp;

             while ( (szTemp = bis.readLine()) != null) {
                szContent+=szTemp+"\n";
             }
             return szContent;
         }finally {
            if(bis!=null){
                bis.close();
            }
        }
    }

    public static String readFile(File file) throws IOException {
        BufferedReader bis=null;
        try {
            bis = new BufferedReader(new InputStreamReader(new FileInputStream( file), ENCODE) );
            String szContent="";
            String szTemp;

            while ( (szTemp = bis.readLine()) != null) {
                szContent+=szTemp+"\n";
            }
            return szContent;
        }finally {
            if(bis!=null){
                bis.close();
            }
        }
    }

    public static String readFile(InputStream inputStream) throws IOException {
        BufferedReader bis=null;
        try {
            bis = new BufferedReader(new InputStreamReader(inputStream,ENCODE) );
            String szContent="";
            String szTemp;

            while ( (szTemp = bis.readLine()) != null) {
                szContent+=szTemp+"\n";
            }
            return szContent;
        }finally {
            if(bis!=null){
                bis.close();
            }
        }
    }


}
