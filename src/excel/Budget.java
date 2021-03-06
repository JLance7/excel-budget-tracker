package excel;

//Joshua Lanctot
//This is a program that uses the Apache POI library to ask the get user input for expenses they will enter and be put into
//an excel file on their computer.

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.Date;
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
            workbook = createNewExcelFile();
        }

        boolean end = false;
        while (!end){
            int option = chooseOption();
            end = useInput(option, workbook, end);
        }


        //Writing created excel file
        FileOutputStream out = new FileOutputStream(new File(userHomeFolder, "Desktop\\Budget Results.xlsx"));
        workbook.write(out);
        workbook.close();
        System.out.println("Excel file saved");
    }

    public static XSSFWorkbook createNewExcelFile(){
        //Create a workbook
        workbook = new XSSFWorkbook();
        //create a worksheet inside of workbook
        sheet = workbook.createSheet("Budget Log");
        //creating cell style
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        XSSFFont font= workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        //create a row object
        row = sheet.createRow(0);

        //create cells and set values
        Cell cell0 = row.createCell(0);
        Cell cell1 = row.createCell(1);
        Cell cell2 = row.createCell(2);
        Cell cell3 = row.createCell(3);
        Cell cell4 = row.createCell(4);
        row.setRowStyle(style);
        cell0.setCellValue("Date");
        cell1.setCellValue("Item/Expense");
        cell2.setCellValue("Money Spent ($)");
        cell3.setCellValue("Total Spent ($)");
        cell4.setCellValue("Budget");
        row.setRowStyle(style);
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
        font.setBold(false);
        style.setFont(font);
        row.setRowStyle(style);
        //auto style columns
        for (int i=0; i< 5; i++){
            sheet.autoSizeColumn(i);
        }
        return workbook;
    }

    //display menu for input
    public static void displayMenu(){
        System.out.println("\n1: Enter expenses");
        System.out.println("2: View Budget");
        System.out.println("3: Change budget");
        System.out.println("4: Save changes and exit");
    }

    //perform actions when each input is entered
    public static int chooseOption(){
        boolean end = false;
        int answer = 0;
        while (answer > 4 || answer < 1){
            try{
                displayMenu();
                System.out.println("Please enter a number 1-4\n");
                answer = input.nextInt();
            } catch (Exception e){
                System.out.println("Please enter a number 1-4");
                input.next();
            }
        }
        return answer;
    }

    //perform actions based on users input
    public static boolean useInput(int answer, XSSFWorkbook workbook, boolean end){
        Scanner input = new Scanner(System.in);
        DateFormat dateFormat = new SimpleDateFormat("MM-dd-yyyy");
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        int i = 1;
        int j = 0;
        switch (answer){
            case 1:
                System.out.println("Enter your item/expense");
                String expense = input.nextLine();
                System.out.println("Enter the cost as a number");
                double cost = input.nextDouble();
                if (workbook.getSheetAt(0).getRow(1) != null && workbook.getSheetAt(0).getRow(1).getCell(0).getCellType() == CellType.BLANK){   //handle first row
                    Cell newDate = workbook.getSheetAt(0).getRow(i).createCell(j);
                    Cell newItem = workbook.getSheetAt(0).getRow(i).createCell(j+1);
                    Cell newCost = workbook.getSheetAt(0).getRow(i).createCell(j+2);
                    newDate.setCellValue(dateFormat.format(new Date()));                      //place current date in new row for expense
                    newItem.setCellValue(expense);
                    newCost.setCellValue(cost);                         //enter new values
                    newDate.setCellStyle(style);                  //style new cells
                    newItem.setCellStyle(style);
                    newCost.setCellStyle(style);
                    for (int k=0; k< 5; k++){
                        sheet.autoSizeColumn(k);
                    }
                    break;
                }

                while (workbook.getSheetAt(0).getRow(i) != null){
                    i++;
                }
                row = sheet.createRow(i);
                for (int m=0; m< 5; m++){
                    sheet.autoSizeColumn(m);
                }
                Cell newDate = workbook.getSheetAt(0).getRow(i).createCell(j);
                Cell newItem = workbook.getSheetAt(0).getRow(i).createCell(j+1);
                Cell newCost = workbook.getSheetAt(0).getRow(i).createCell(j+2);

                newDate.setCellValue(dateFormat.format(new Date()));                      //place current date in new row for expense
                newItem.setCellValue(expense);
                newCost.setCellValue(cost);                         //enter new values

                newDate.setCellStyle(style);                  //style new cells
                newItem.setCellStyle(style);
                newCost.setCellStyle(style);

                int z = 0;
                while (workbook.getSheetAt(0).getRow(z) != null){
                    z++;
                }
                Integer p = z;
                String strFormula= "SUM(C2:C" + p.toString() + ")";
                workbook.getSheetAt(0).getRow(1).getCell(3).setCellFormula(strFormula);
                row = sheet.getRow(1);
                row.setRowStyle(style);
                for (int k=0; k< 5; k++){
                    sheet.autoSizeColumn(k);
                }
                end = false;
                break;
            case 2:
                z = 0;
                while (workbook.getSheetAt(0).getRow(z) != null){
                    z++;
                }
                p = z;
                strFormula= "SUM(C2:C" + p.toString() + ")";
                workbook.getSheetAt(0).getRow(1).getCell(3).setCellFormula(strFormula);
                System.out.println("Your budget is: $" + workbook.getSheetAt(0).getRow(1).getCell(4));
                if (workbook.getSheetAt(0).getRow(1).getCell(4).getCellType() == CellType.BLANK)
                    System.out.println("You have currently have zero expenses.");
                else{
                    double budget = (workbook.getSheetAt(0).getRow(1).getCell(4).getNumericCellValue());
                    Cell cell = (workbook.getSheetAt(0).getRow(1).getCell(3));
                    FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                    CellValue cellValue = evaluator.evaluate(cell);
                    double total = cellValue.getNumberValue();
                    double difference = budget - total;
                    System.out.println("Your total expenses is: $" + cellValue.getNumberValue());
                    DecimalFormat decimalFormat = new DecimalFormat("#.00");
                    if (difference >= 0){
                        System.out.println("You have $" + decimalFormat.format(difference) + " money left for your budget.");
                    }
                    else{
                        double money = -1 * difference;
                        System.out.println("You are $" + decimalFormat.format(money) + " over your budget.");
                    }
                }
                end = false;
                break;
            case 3:
                System.out.println("What would you like your new budget to be? ");
                double newBudget = input.nextDouble();
                workbook.getSheetAt(0).getRow(1).getCell(4).setCellValue(newBudget);
                System.out.println("Your budget has been changed to: " + newBudget);
                end = false;
                break;
            case 4:
                end = true;
                break;
        }
        return end;
    }
}
