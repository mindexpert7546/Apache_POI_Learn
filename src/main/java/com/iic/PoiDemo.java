package com.iic;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class PoiDemo {
    public static void main(String[] args) throws IOException {
        String filePath = "C:\\Users\\Kundan.kumar\\IdeaProjects\\Apache_POI_\\mypoi.xlsx";
        File myFile = new File(filePath);
        Workbook workbook = new XSSFWorkbook();
        FileOutputStream fileOutputStream = new FileOutputStream(myFile);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("File Created..");
    }
}