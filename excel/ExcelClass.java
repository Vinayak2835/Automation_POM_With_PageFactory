package com.salenium.excel;

import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelClass {
	
	 String name = "";
	 String password = "";
	 //int password = 0;
	    
public static void main(String[] args)  {
	    ExcelClass excelClass = new ExcelClass();
		
		try {
			
		   String excel =".\\excel\\data2.xlsx";
		
		   FileInputStream fileInputStream = new FileInputStream(excel);
		   XSSFWorkbook workBook = new XSSFWorkbook(fileInputStream);
		   
		  // XSSFSheet sheet = workBook.getSheet("Sheet1");
		   XSSFSheet sheet = workBook.getSheetAt(0);
		   int totalNumberOfRows = sheet.getPhysicalNumberOfRows();
		  System.out.println("Total Number of Rows "+totalNumberOfRows);
		   
		   int rows = sheet.getLastRowNum();
		   int coln = sheet.getRow(1).getLastCellNum();
		   
		   Iterator itr = sheet.iterator();
		   
		   while(itr.hasNext()) {
			   XSSFRow row = (XSSFRow)itr.next();
			   Iterator cellIterator = row.cellIterator();
			   
			   while(cellIterator.hasNext()) {
				   XSSFCell  cell =(XSSFCell)cellIterator.next();
				   
				  switch(cell.getCellType()) {
				  case 1:
					     if(cell.getStringCellValue().equalsIgnoreCase("Admin")) {
					        excelClass.name = cell.getStringCellValue()+" ";
				            System.out.println("UserName is: "+excelClass.name);
				         }
				         if(cell.getStringCellValue().equalsIgnoreCase("admin123")) {
				        	 excelClass.password = cell.getStringCellValue();
				        	 System.out.println("yes matches and password is "+excelClass.password);
				         }
  		                 break;
			     case 2: 
				     int number = (int)cell.getNumericCellValue();
				     System.out.println(number);
			         break;
				  
				   }
				  
				}
		   }
		   
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
}
