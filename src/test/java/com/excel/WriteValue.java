package com.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteValue {
	
	public static void main(String[] args)throws Throwable {
		File file=new File(System.getProperty("user.dir")+"\\src\\test\\resources\\sample1.xlsx");
		FileOutputStream stream=new FileOutputStream(file);
		Scanner scan=new Scanner(System.in);
		System.out.println("Enter Row Count: ");
		int rowCount = scan.nextInt();
		System.out.println("Enter Cell Count: ");
		int cellCount = scan.nextInt();
		Workbook workbook=new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Vasanth");
		
		for(int i=0;i<rowCount;i++) {
			Row row = sheet.createRow(i);
			for(int j=0;j<cellCount;j++) {
				Cell cell = row.createCell(j);
				String value = scan.next();
				cell.setCellValue(value);
			}
		}
		scan.close();
		workbook.write(stream);
		stream.close();
		workbook.close();
		
		System.out.println("Success");
	}

}
