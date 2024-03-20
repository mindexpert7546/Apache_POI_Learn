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

public class PoiDemo_table {
    public static void main(String[] args) throws IOException {
        String filePath = "C:\\Users\\Kundan.kumar\\Documents\\Apache_POI_\\OutPut\\mypoiTable.xlsx"; // File path
        File myFile = new File(filePath); //To create the file

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet1 = workbook.createSheet("Sheet1");

        // 2D array
        String[][] tableData = {
                {"No", "Name", "Position", "Salary"},
                {"1", "Krishna", "Intern", "15000"},
                {"2", "Aarav", "Junior Developer", "25000"},
                {"3", "Advik", "Senior Developer", "35000"},
                {"4", "Ishaan", "Manager", "45000"},
                {"5", "Vihaan", "Product Manager", "55000"},
                {"6", "Sai", "Quality Assurance", "30000"},
                {"7", "Arjun", "System Analyst", "40000"},
                {"8", "Anika", "HR Executive", "30000"},
                {"9", "Ved", "Finance Manager", "50000"},
                {"10", "Myra", "Marketing Coordinator", "35000"},
                {"11", "Ayaan", "Operations Manager", "45000"},
                {"12", "Ananya", "Chief Technology Officer", "60000"},
                {"13", "Kabir", "Software Engineer", "38000"},
                {"14", "Navya", "UX/UI Designer", "32000"},
                {"15", "Aarush", "Data Analyst", "34000"},
                {"16", "Priya", "Content Writer", "29000"},
                {"17", "Rohan", "SEO Specialist", "31000"},
                {"18", "Aditi", "Social Media Manager", "33000"},
                {"19", "Karan", "Product Designer", "41000"},
                {"20", "Ira", "Business Analyst", "37000"},
                {"21", "Siddharth", "DevOps Engineer", "47000"},
                {"22", "Mira", "Project Manager", "43000"},
        };

        Row row = null;
        Cell cell = null;
        int rowPosition = 4;
        int colPosition = 5;

        //Insert the data into the excel
        for(int i=0; i<tableData.length; i++){
           row = sheet1.createRow(rowPosition+i);
           for(int j=0; j<tableData[0].length; j++){
               cell=row.createCell(colPosition+j);
               if(i!=0) {
                   if (j == 0 || j == 3) {
                       cell.setCellValue(Integer.parseInt(tableData[i][j]));
                   } else {
                       cell.setCellValue(tableData[i][j]);
                   }
               } else {
                   cell.setCellValue(tableData[i][j]);
               }
           }
        }


        FileOutputStream fout = new FileOutputStream(myFile);
        workbook.write(fout);
        fout.close();
        System.out.println("File created successfully...");
    }
}
