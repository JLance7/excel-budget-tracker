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
    static Scanner input = new Scanner(System.in);
    static XSSFWorkbook workbook;
    static XSSFSheet sheet;
    static XSSFRow row;

    public static void main(String[] args) throws IOException {
        String userHomeFolder = System.getProperty("user.home");
        //Get existing file to be read
        File file = new File(userHomeFolder, "Desktop\\Budget Results.xlsx");
        //if the file Budget Results already exists on the desktop open it, if not create and start file
        if (file.isFile()){
            FileInputStream fip = new FileInputStream(file);
            workbook = new XSSFWorkbook(fip);
            sheet = workbook.getSheetAt(0);
        }
        else {
            createNewExcelFile();
        }

        int option = chooseOption();
        useInput(option);

        //Writing created excel file
        FileOutputStream out = new FileOutputStream(new File(userHomeFolder, "Desktop\\Budget Results.xlsx"));
        workbook.write(out);
        workbook.close();
        System.out.println("Excel file saved");
    }

    public static void createNewExcelFile(){
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

        cell0.setCellValue("Date");
        cell1.setCellValue("Item/Expense");
        cell2.setCellValue("Money Spent ($)");
        cell3.setCellValue("Total Spent");
        cell4.setCellValue("Budget");

        System.out.println("What is your budget?");     //set budget cell
        double budget = input.nextDouble();
        row = sheet.createRow(1);
        Cell cell5 = row.createCell(0);
        Cell cell6 = row.createCell(1);
        Cell cell7 = row.createCell(2);
        Cell cell8 = row.createCell(3);
        Cell cell9 = row.createCell(4);
        cell9.setCellValue(budget);

        //style cells
        row.setRowStyle(style);
        //auto style columns
        for (int i=0; i< 5; i++){
            sheet.autoSizeColumn(i);
        }
    }

    //display menu for input
    public static void displayMenu(){
        System.out.println("Enter a number 1-3 for which action you would like to perform.");
        System.out.println("1: Enter expenses");
        System.out.println("2: View Budget");
        System.out.println("3: Change budget");

    }

    //perform actions when each input is entered
    public static int chooseOption(){
        boolean end = false;
        int answer = 0;
        while (answer <= 3 && answer >= 1){
            displayMenu();
            System.out.println("Please enter a number 1-3\n");
            answer = input.nextInt();
        }
        return answer;
    }

    //perform actions based on users input
    public static void useInput(int answer){
        Scanner input = new Scanner(System.in);
        switch (answer){
            case 1:
                System.out.println("Enter your item/expense");
                String expense = input.nextLine();
                break;
            case 2:
                System.out.println("Your budget is: " + workbook.getSheetAt(0).getRow(1).getCell(9);
            case 3:
                break;
        }
    }
}
