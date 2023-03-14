package com.root.ExcelProject;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelClass {
	
	public static void main(String[] args) {
		
		try {
			
			FileInputStream file = new FileInputStream(new File("Employee_info_report.xlsx"));
			
			XSSFWorkbook wb = new XSSFWorkbook(file);
			
			XSSFSheet sh = wb.getSheetAt(0);
			
			Iterator<Row> iterator = sh.iterator();
			
			while(iterator.hasNext())
			{
				Row row = iterator.next();
				
				Iterator<Cell> cellIterator = row.cellIterator();
				
				while(cellIterator.hasNext())
				{
					Cell cell= cellIterator.next();
					
					System.out.print(cell.getStringCellValue()+" ");
				}
				System.out.println("");
			}
			file.close();
			wb.close();
			
			
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		
	}

}
