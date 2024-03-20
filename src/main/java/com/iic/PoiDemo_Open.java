package com.iic;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class PoiDemo_Open {
    public static void main(String[] args) throws IOException {
        String filePath = "C:\\Users\\Kundan.kumar\\Documents\\Apache_POI_\\OutPut\\mypoi.xlsx"; // File path
        File myFile = new File(filePath); //To create the file

        FileInputStream fis = null;
        fis = new FileInputStream(myFile);

        if(myFile.isFile() && myFile.exists()){
            System.out.println("File open successfully...");
        } else{
            System.out.println("File not open.");
        }
        fis.close();
    }
}
