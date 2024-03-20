package com.iic;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class PoiDemo_Font {
    public static void main(String[] args) throws IOException {
        String filePath = "C:\\Users\\Kundan.kumar\\Documents\\Apache_POI_\\OutPut\\mypoiFont.xlsx"; // File path
        File myFile = new File(filePath); //To create the file

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet1 = workbook.createSheet("Sheet1");

        CellStyle cellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.TOP);

        font.setBold(true);
        font.setItalic(true);

        Row row = sheet1.createRow(1);
        Cell cell = row.createCell(3);
        cell.setCellValue("Krishna");

        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle);


        FileOutputStream fout = new FileOutputStream(myFile);
        workbook.write(fout);

        fout.close();
        System.out.println("File created successfully...");
    }
}
