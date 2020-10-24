/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package readexcelfile;
import java.io.File;  
import java.io.FileInputStream;  
import java.io.IOException;  
import org.apache.poi.hssf.usermodel.HSSFSheet;  
import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.FormulaEvaluator;  
import org.apache.poi.ss.usermodel.Row;  

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.Scanner;
import java.util.*;
/**
 *
 * @author DELL
 */
public class ReadExcelFile {
static FileInputStream inputStream;
static Workbook workbook;
static Sheet firstSheet;
static Iterator<Row> iterator;
static List<Data> data = new ArrayList<Data>();  
    public static void main(String[] args) throws IOException {
 // Scanner scanner = new Scanner(System.in);
  
//String columnToRead = scanner.nextLine();
//        String columnsToRead =scanner.nextLine();
         inputStream = new FileInputStream(new File("C:\\Users\\DELL\\Desktop\\SRS\\myExcel.xlsx"));
         
        workbook = new XSSFWorkbook(inputStream);
        firstSheet = workbook.getSheetAt(0);
         iterator = firstSheet.iterator();
         
         
         readColumn("Status");


//addToList2();




//String[] columns = columnsToRead.split(",");

for(int i=0; i<data.size(); i++){
    //System.out.println(columns[i]);
//    readColumn(columns[i]);

//    System.out.print(data.get(i).name+"\t");
//    System.out.print(data.get(i).marks+"\t");
//    System.out.print(data.get(i).status+"\t");
        System.out.print(data.get(i).toString()+"\t");

    System.out.println("");
}




      
      //      Row nextRow = iterator.next();
       //     Iterator<Cell> cellIterator = nextRow.cellIterator();
       /*     
   while(cellIterator.hasNext())
            {           Cell cell = cellIterator.next();
                 
                
               String columnName = cell.getStringCellValue();
      
              // System.out.println("columnName is: "+columnName+" columnToRead: "+columnToRead);
               if (columnName.equals(columnToRead)){
                 
                   System.out.println("found: "+columnName);
                 readColumn(columnToRead,nextRow);
               }
    }*/
        workbook.close();
        inputStream.close();
    

    
    }
    
    public static void readColumn(String columnToRead){
      //System.out.println("Reading Column: "+columnName);
   boolean columnExists =false;
   int cellIndex=0;
Data data = new Data();
   
iterator = firstSheet.iterator();
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            while(cellIterator.hasNext()){
            Cell cell = cellIterator.next();
                //System.out.println(cell.getStringCellValue());
                String currentColumn  = cell.getStringCellValue();
                if(currentColumn.equals(columnToRead)){
                    //System.out.println("found");
                    
                    if(currentColumn.equals("Name"))
                        cellIndex=0;
                    else if(currentColumn.equals("Marks"))
                        cellIndex=1;
                    else if(currentColumn.equals("Status"))
                        cellIndex=2;
                    
                    columnExists = true;
                   // System.out.println();
                    
                    //System.out.println(nextRow.getCell(0).getStringCellValue());
                    
                    
                }
                
            }
            //iterator=null;
            //iterator= firstSheet.iterator();
            Row row;
            while(iterator.hasNext()){

                row = iterator.next();
                 //row = iterator.next();
                 //System.out.println(row.getCell(cellIndex).getStringCellValue());
                 Cell cell = row.getCell(cellIndex);
                 switch (cell.getCellType()){
                        case Cell.CELL_TYPE_STRING:
                        System.out.print(cell.getStringCellValue());
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        System.out.print(cell.getBooleanCellValue());
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print(cell.getNumericCellValue());
                        break;
                 
                 }
                 System.out.println("");
                
                
            }
            System.out.println("====================================");
            if(!columnExists){
                System.out.println("the given column does not exist");
            }
            
      
      
      
    }
    public static void addToList2(){
    iterator = firstSheet.iterator();
    Row currentRow = iterator.next();
    Data dataObj ;
       //currentRow = iterator.next();
    while(iterator.hasNext()){
           currentRow= iterator.next();
    //Iterator<Cell> cellIterator = currentRow.cellIterator();
          // while (cellIterator.hasNext()) {
      //          Cell cell = cellIterator.next();
                 
                String name = currentRow.getCell(0).getStringCellValue();
                double marks = currentRow.getCell(1).getNumericCellValue();
                String status = currentRow.getCell(2).getStringCellValue();
                
               dataObj = new Data(name,marks,status);
               data.add(dataObj);
               // System.out.println("added");
            
                //System.out.print(" - ");
            //}
    
    
    
    }
        
    }
    
    
    public static void addToList() throws IOException{
     Data data = new Data();
                String excelFilePath = "C:\\Users\\DELL\\Desktop\\SRS\\myExcel.xlsx";
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
         
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = firstSheet.iterator();
         
        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
             
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                 
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        System.out.print(cell.getStringCellValue());
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        System.out.print(cell.getBooleanCellValue());
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print(cell.getNumericCellValue());
                        break;
                }
                System.out.print(" - ");
            }
            System.out.println();
        }
         
        workbook.close();
        inputStream.close();
        
    }
    
    }


class Data{

public String name;
public double marks;
public String status;

public Data(){

}

    @Override
    public String toString() {
        return  "name=" + name + ", marks=" + marks + ", status=" + status;
    }

    public String getName() {
        return name;
    }

    public double getMarks() {
        return marks;
    }

    public String getStatus() {
        return status;
    }

    public Data(String name, double marks, String status) {
        this.name = name;
        this.marks = marks;
        this.status = status;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setMarks(double marks) {
        this.marks = marks;
    }

    public void setStatus(String status) {
        this.status = status;
    }
    



}
