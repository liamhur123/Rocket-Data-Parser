

package utils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.awt.*;
import java.awt.event.*;


import javax.swing.*;
import javax.swing.filechooser.FileSystemView;
import java.io.FileOutputStream;
import java.util.*;

import java.io.IOException;

public class ExcelUtils implements ActionListener{
 static XSSFWorkbook workbook;
 static XSSFSheet sheet;
 static TreeMap<String, HashSet<String>> treeMap;



 public ExcelUtils(String excelPath){
  
  try {
   workbook = new XSSFWorkbook(excelPath);
   sheet = workbook.getSheet("Sheet1");
  }catch(Exception exp){
  System.out.println(exp.getCause());
  System.out.println(exp.getMessage());
  exp.printStackTrace();
  }
 }
 
 
 public void main(String[] args) throws IOException {
 }
 public static int getRowCount(){
  try {
   int rowCount = sheet.getPhysicalNumberOfRows(); //gets number of rows used in Excel sheet
   return rowCount; //returns the integer value of rows in the Excel sheet

  }catch(Exception exp){
   System.out.println(exp.getCause());
   System.out.println(exp.getMessage());
   exp.printStackTrace();
  }
  return 0;
 }
 public static String getCellData(int rowNum, int colNum) throws IOException {

  return (String) sheet.getRow(rowNum).getCell(colNum).getStringCellValue(); //Gets data within specified cell and row number

 }
public static TreeMap<String,HashSet<String>> classesWithNames() throws IOException {
    treeMap = new TreeMap<String,HashSet<String>>(); //generates storage and nested storage for class names/times and Names of participants in that class
    int rowCount = getRowCount(); // gets the number of rows
    for (int i = 1; i < rowCount; i++) {
        if (!treeMap.containsKey(getCellData(i,0))) { // if map doesn't have the class
            treeMap.put(getCellData(i,0),new HashSet<String>()); // add the class to the map
            treeMap.get(getCellData(i,0)).add(getCellData(i,25).replace(",","") ); // also adds the name to the map
        }else{
            treeMap.get(getCellData(i,0)).add(getCellData(i,25).replace(",","") ); // if it already has class time, just updates names to the map
        }
    }

    return treeMap;
}

    public static void WriteClassNamesandTimes() throws IOException {
            classesWithNames();
            XSSFWorkbook wb = new XSSFWorkbook(); // creates workbook
            XSSFSheet sheet = wb.createSheet("Roster Names"); // creates/names the sheet in the workbook
            int count = 0; // counter
        for(String Classes : treeMap.keySet()){ // for each class in the map
            XSSFRow classRow = sheet.createRow(count); // create a new row
                XSSFCell Cell = classRow.createCell(0); //creat a new cell
                Cell.setCellValue((String) Classes); //set that cell to the class name
                count++; // increase count
                for(String Names : treeMap.get(Classes)){ // for each name in the map
                    XSSFRow nameRow = sheet.createRow(count); // create a new row
                    XSSFCell lastName = nameRow.createCell(0); // create a new cell for last name
                    XSSFCell firstName = nameRow.createCell(1); // create a new cell for first name
                    String[] nameParts = Names.split(" "); // split the name into parts
                    if(nameParts.length == 2){ // if the name is only first and last
                        lastName.setCellValue(nameParts[0]); // set first cell to last name
                        firstName.setCellValue(nameParts[1]); // set the second cell to the first name
                    }else if(nameParts.length > 2){ // if the name has a middle name or more
                        lastName.setCellValue(nameParts[0] + " " +nameParts[2]); // set first cell to last name
                        firstName.setCellValue(nameParts[1]); // set second cell to first name
                    }
                    count++;
                   }




        }
       /*
       ----- to get the file to output on the desktop -----
       FileSystemView fliesys = FileSystemView.getFileSystemView();
       fliesys.getHomeDirectory();     // when in path you need to add "\\folderName\\fileName.xlsx"

        */
        String filePath = "D:\\Projects\\test\\test.xlsx"; // sets file path
        //String filePath = "C:\\Users\\liamh\\OneDrive\\Desktop\\test\\test.xlsx";
            FileOutputStream outstream = new FileOutputStream(filePath); // creates the output stream to file location
            wb.write(outstream); //writes our data to the Excel workbook
            outstream.close(); // ends the file stream
            System.out.println("test.xlsx success"); // notification that file has successfully been created and written
        }
    public static String wantDate(){
     Scanner ask = new Scanner(System.in);
     System.out.println("Registrations after date(month day, year): ");
     String date = ask.next();
     return date;
    }
    public static void registrationDate() throws IOException {
        classesWithNames();
        String date = wantDate();
        XSSFWorkbook wb = new XSSFWorkbook(); // creates workbook
        XSSFSheet sheet = wb.createSheet("Roster Names"); // creates/names the sheet in the workbook
        int count = 0; // counter
        for(String Classes : treeMap.keySet()){ // for each class in the map
            XSSFRow classRow = sheet.createRow(count); // create a new row
            XSSFCell Cell = classRow.createCell(0); //creat a new cell
            Cell.setCellValue((String) Classes); //set that cell to the class name
            count++; // increase count
            for(String Names : treeMap.get(Classes)) { // for each name in the map

                // need to make it so that it gets all the names after a certain date
                // also need to go through each name/class and "assign" date to them cause this is not enough for it to work

                if (getCellData(1,31) == date ) { // check date to see if correct


                    XSSFRow nameRow = sheet.createRow(count); // create a new row
                    XSSFCell lastName = nameRow.createCell(0); // create a new cell for last name
                    XSSFCell firstName = nameRow.createCell(1); // create a new cell for first name
                    String[] nameParts = Names.split(" "); // split the name into parts
                    if (nameParts.length == 2) { // if the name is only first and last
                        lastName.setCellValue(nameParts[0]); // set first cell to last name
                        firstName.setCellValue(nameParts[1]); // set the second cell to the first name
                    } else if (nameParts.length > 2) { // if the name has a middle name or more
                        lastName.setCellValue(nameParts[0] + " " + nameParts[2]); // set first cell to last name
                        firstName.setCellValue(nameParts[1]); // set second cell to first name
                    }
                    count++;
                }



            }


        }
/*
       ----- to get the file to output on the desktop -----
       FileSystemView fliesys = FileSystemView.getFileSystemView();
       fliesys.getHomeDirectory();     // when in path you need to add "\\folderName\\fileName.xlsx"


 */

        String filePath = "D:\\Projects\\test\\test.xlsx"; // sets file path
        FileOutputStream outstream = new FileOutputStream(filePath); // creates the output stream to file location
        wb.write(outstream); //writes our data to the Excel workbook
        outstream.close(); // ends the file stream
        System.out.println("test.xlsx success"); // notification that file has successfully been created and written
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        throw new UnsupportedOperationException("Not supported yet."); // Generated from nbfs://nbhost/SystemFileSystem/Templates/Classes/Code/GeneratedMethodBody
    }




}


/*
----------------------------------------------------------------------------------------------------------------------------------------
    public static void getAllNames() throws IOException {

        int rowCount = getRowCount(); // grabs the number of rows
        for (int i = 1; i < rowCount; i++) {
            System.out.println(getCellData(i,25).replace(",","")); //prints all names without comas - lastName firstName
        }

    }
    public static void getClassDays() throws IOException {
        HashSet<String> classes = new HashSet<String>(); //generates storage
        int rowCount = getRowCount(); // grabs the row count
        for (int i = 1; i < rowCount; i++) {
            classes.add(getCellData(i,0)); //for each row it grabs the first column to get class names and times
        }
        System.out.println(classes);
    }
----------------------------------------------------------------------------------------------------------------------------------------



TODO:
Starter mini sessions

Names after a certain register date

skate rental charge

NOTES:
Cell 0 will always be activity name ~ "Basic 1 (Mon 5:45 PM ) - 7419"
Cell 4 will start time ~ "5:45 PM"
Cell 6 will be day of the week ~ "M" "Tu" "W" "Th" "F" "Sa" "Su"
Cell 8 will be Activity Category ~ "Ice Skating" "Learn to Play Hockey"
Cell 25 will be names ~ "Smith, John"
Cell 31 is Date of registration






https://www.youtube.com/watch?v=ipjl49Hgsg8&list=PLUDwpEzHYYLsN1kpIjOyYW6j_GLgOyA07
 */