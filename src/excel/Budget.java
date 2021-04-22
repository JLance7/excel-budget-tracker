package excel;

//Joshua Lanctot
//This is a program that uses the Apache POI library to ask the get user input for expenses they will enter and be put into
//an excel file on their computer.

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Scanner;

public class Budget {
    public static void main(String[] args) throws IOException {
        Scanner input = new Scanner(System.in);
        String userHomeFolder = System.getProperty("user.home");
        //Get existing file to be read
        File file = new File(userHomeFolder, "Desktop\\Budget Results.xlsx");
        XSSFWorkbook workbook;
        XSSFSheet sheet;
        XSSFRow row;

        //if the file Budget Results already exists on the desktop ask for user input, if not setup first row and then ask for input
        if (file.isFile()){
            FileInputStream fip = new FileInputStream(file);

            workbook = new XSSFWorkbook(fip);
            sheet = workbook.getSheetAt(0);

        }
        else{
            //Create a workbook
            workbook = new XSSFWorkbook();

            //create a worksheet inside of workbook
            sheet = workbook.createSheet("Budget Log");

            //creating cell style
            CellStyle style = workbook.createCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setVerticalAlignment(VerticalAlignment.CENTER);

            //create a row object
            row = sheet.createRow(0);

            //create cells and set values
            Cell cell0 = row.createCell(0);
            Cell cell1 = row.createCell(1);
            Cell cell2 = row.createCell(2);
            Cell cell3 = row.createCell(3);
            Cell cell4 = row.createCell(4);

            //style cells
            cell0.setCellStyle(style);
            cell1.setCellStyle(style);
            cell2.setCellStyle(style);
            cell3.setCellStyle(style)
            cell4.setCellStyle(style);

            //auto style columns
            for (int i=0; i< 5; i++){
                sheet.autoSizeColumn(i);
            }

            cell0.setCellValue("Date");
            cell1.setCellValue("Item/Expense");
            cell2.setCellValue("Money Spent ($)");
            cell3.setCellValue("Total Spent");
            cell4.setCellValue("Budget");
        }

        //perform actions when each input is entered
        boolean end = false;
        int answer;
        while (!end){
            displayMenu();
            answer = input.nextInt();
            while (answer > 3 && answer < 1){
                System.out.println("Please enter a number 1-3\n");
                displayMenu();
                answer = input.nextInt();
            }
            useInput(answer);
        }


        //Writing created excel file
        FileOutputStream out = new FileOutputStream(new File(userHomeFolder, "Desktop\\Budget Results.xlsx"));
        workbook.write(out);
        workbook.close();
        System.out.println("Excel file saved");
    }

    //display menu for input
    public static void displayMenu(){
        System.out.println("Enter a number 1-3 for which action you would like to perform.");
        System.out.println("1: Enter expenses");
        System.out.println("2: View Budget");
        System.out.println("3: Change budget");

    }

    //perform actions based on users input
    public static void useInput(int answer){
        switch (answer){
            case 1:
                
        }
    }
}
