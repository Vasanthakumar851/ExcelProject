package com.excel;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadValue {
	
	public static void main(String[] args)throws Throwable {
		File file=new File(System.getProperty("user.dir")+"\\src\\test\\resources\\ExcelDatas\\sample.xlsx");
		FileInputStream stream=new FileInputStream(file);
		Workbook workbook=new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet("Sheet1");
		int rowCount = sheet.getPhysicalNumberOfRows();
		
		for(int i=0;i<rowCount;i++) {
			System.out.println();
			Row row = sheet.getRow(i);
			int cellCount = row.getPhysicalNumberOfCells();
			for(int j=0;j<cellCount;j++) {
				Cell cell = row.getCell(j);
				System.out.print(cell+"\t\t\t");
				
			}
		}
		stream.close();
		workbook.close();
		
	}

}
