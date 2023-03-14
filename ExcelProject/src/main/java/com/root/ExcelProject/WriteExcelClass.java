package com.root.ExcelProject;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

@Component
public class WriteExcelClass {
	
	public static void main(String[] args) {
		
		XSSFWorkbook wb= new XSSFWorkbook();
		
		XSSFSheet sh= wb.createSheet("Employee_Information");
		
		List<Map<String,Object>> records = new ArrayList<Map<String,Object>>();
		
		Map<String,Object> map1 = new TreeMap<String,Object>();
		map1.put("employeeNo", String.valueOf(001));
		map1.put("employeeName", String.valueOf("vinay"));
		map1.put("designation", String.valueOf("SDE1"));
		map1.put("location", String.valueOf("Hyderabad"));
		
		Map<String,Object> map2 = new TreeMap<String,Object>();
		map2.put("employeeNo", String.valueOf(002));
		map2.put("employeeName", String.valueOf("Ram"));
		map2.put("designation", String.valueOf("SDE2"));
		map2.put("location", String.valueOf("Bangalore"));
		
		records.add(map1);
		records.add(map2);
		
		reportCaption(wb, sh);
		reportHeadings(wb,sh);
		
		int rowIndex=1;
		for(Map<String,Object> map : records) {
			XSSFRow row = sh.createRow(++rowIndex);
			XSSFCell cell = row.createCell(0);
			cell.setCellValue(String.valueOf(map.get("employeeNo")));
			cell = row.createCell(1);
			cell.setCellValue(String.valueOf(map.get("employeeName")));
			cell = row.createCell(2);
			cell.setCellValue(String.valueOf(map.get("designation")));
			cell = row.createCell(3);
			cell.setCellValue(String.valueOf(map.get("location")));
		}
		try {
			FileOutputStream fos = new FileOutputStream(new File("Employee_Info_report.xlsx"));
			wb.write(fos);
			fos.close();
			System.out.println("Export successfull.");
		}
		catch(Exception ex) {
			System.out.println(ex.getMessage());
		}
		
	}
	
	private static void reportCaption(XSSFWorkbook wb, XSSFSheet sheet) {
		XSSFRow row= sheet.createRow(0);
		XSSFCell cell = row.createCell(0);
		cell.setCellValue("Employees Information Report");
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));
	}
	
	private static void reportHeadings(XSSFWorkbook wb, XSSFSheet sheet) {
		XSSFRow row = sheet.createRow(1);
		XSSFCell cell = row.createCell(0);
		cell.setCellValue("Employee No");
		cell = row.createCell(1);
		cell.setCellValue("Employee Name");
		cell = row.createCell(2);
		cell.setCellValue("Designation");
		cell = row.createCell(3);
		cell.setCellValue("Location");
		
	}
}
