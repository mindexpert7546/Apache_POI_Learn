package com.iic;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class PoiDemo {
    public static void main(String[] args) throws IOException {
        String filePath = "C:\\Users\\Kundan.kumar\\Documents\\Apache_POI_\\OutPut\\mypoi.xlsx"; // File path
        File myFile = new File(filePath); //To create the file

        Workbook workbook = new XSSFWorkbook(); // code for create the excel file
        Sheet sheet1 = workbook.createSheet("sheet1"); // create the new sheet in mypoi.xlsx

        FileOutputStream fileOutputStream = new FileOutputStream(myFile);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("File Created..");
    }
}