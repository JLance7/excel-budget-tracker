package excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
//Joshua Lanctot
//This is a program that uses the Apache POI library to ask the get user input for expenses they will enter and be put into
//an excel file on their computer.
public class Budget {
    public static void main(String[] args){
        //Create a workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //creating cell style
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        //create a worksheet inside of workbook
        XSSFSheet sheet = workbook.createSheet("Sheet 1");

        //create a row object
        XSSFRow row;

        //create cells and set values
        row = sheet.createRow(0);
        Cell cell0 = row.createCell(0);
        Cell cell1 = row.createCell(1);
        Cell cell2 = row.createCell(2);

        cell0.setCellStyle(style);
        cell1.setCellStyle(style);
        cell2.setCellStyle(style);

        cell0.setCellValue("Date");
        cell1.setCellValue("$   Spent");
        cell2.setCellValue("Total");

        for (int i = 0; i < 10; i++){
            row = sheet.createRow(i+1);

            for (int j = 0; j < 3; j++){
                Cell cell = row.createCell(j);
            }
        }

        //auto style columns
        for (int i=0; i< 5; i++){
            sheet.autoSizeColumn(i);
        }

        //Writing created excel file
        String userHomeFolder = System.getProperty("user.home");
        FileOutputStream out = new FileOutputStream(new File(userHomeFolder, "Desktop\\Results.xlsx"));
        workbook.write(out);
        workbook.close();
        System.out.println("Excel file created");
    }
}
