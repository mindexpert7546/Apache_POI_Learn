package com.iic;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class PoiDemo_BackgroundColor {
    public static void main(String[] args) throws IOException {
        String filePath = "C:\\Users\\Kundan.kumar\\Documents\\Apache_POI_\\OutPut\\mypoiBackGroundColor.xlsx"; // File path
        File myFile = new File(filePath); //To create the file

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet1 = workbook.createSheet("Sheet1");

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
        // Set fill pattern and foreground color
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());

        Row row = sheet1.createRow(1);
        Cell cell = row.createCell(3);
        cell.setCellValue("Krishna");
        cell.setCellStyle(cellStyle);

        FileOutputStream fout = new FileOutputStream(myFile);
        workbook.write(fout);

        fout.close();
        System.out.println("File created successfully...");
    }
}
