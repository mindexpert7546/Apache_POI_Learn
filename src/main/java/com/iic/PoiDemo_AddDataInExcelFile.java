package com.iic;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class PoiDemo_AddDataInExcelFile {
    public static void main(String[] args) throws IOException {
        String filePath = "C:\\Users\\Kundan.kumar\\Documents\\Apache_POI_\\OutPut\\mypoi1.xlsx"; // File path
        File myFile = new File(filePath); //To create the file

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet1 = workbook.createSheet("Sheet1");

        //Create the row object and cell object
        Row row = null;
        Cell cell = null;

        row = sheet1.createRow(2);
        cell = row.createCell(3);
        cell.setCellValue("Krishna");

        row = sheet1.createRow(3);
        cell = row.createCell(4);
        cell.setCellValue("Singh");

        FileOutputStream fout = new FileOutputStream(myFile);
        workbook.write(fout);

        fout.close();
        System.out.println("File created successfully...");
    }
}
