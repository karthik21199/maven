package com.DataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataFirst {
	
	public static void methodOne()throws IOException {
		
		
		File f = new File("C:\\Users\\LENOVO\\Desktop\\Project_11Am\\Book2.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fis);
		
		
		Sheet sheetAt = w.getSheetAt(0);
		int numberOfRows = sheetAt.getPhysicalNumberOfRows();
		
		for (int i = 0; i < numberOfRows; i++) {
			Row row = sheetAt.getRow(i);
			int numberOfCells = row.getPhysicalNumberOfCells();
			for (int j = 0; j < numberOfCells; j++) {
				Cell cell = row.getCell(j);
				CellType cellType = cell.getCellType();
				if (cellType.equals(cellType.STRING)) {
					String Value = cell.getStringCellValue();
					System.out.println(Value);
					
				}
				else if (cellType.equals(cellType.NUMERIC)) {
					double Value = cell.getNumericCellValue();
					int num=(int) Value;
					System.out.println(num);
				}
			}
			
			
			
		}
		       
		System.out.println("------Row Data----");
		Row row1 = sheetAt.getRow(2);
		int numberOfCells = row1.getPhysicalNumberOfCells();
		for (int i = 0; i < numberOfCells; i++) {
		
			Cell cell = row1.getCell(i);
			CellType cellType = cell.getCellType();
			if (cellType.equals(cellType.STRING)) {
				String Value = cell.getStringCellValue();
				System.out.println(Value);
				
			}
			else if (cellType.equals(cellType.NUMERIC)) {
				double Value = (int)cell.getNumericCellValue();
				int num=(int) Value;
				System.out.println(num);				
			}
			
			
			
			
		}
		
	     
		System.out.println("------Particular Data----");
			Row row3 = sheetAt.getRow(0);
			for (int i = 1; i < 2; i++) {
				Cell cell1 = row3.getCell(i);
				CellType cellType1 = cell1.getCellType();
			
			if (cellType1.equals(cellType1.STRING)) {
				String Value = cell1.getStringCellValue();
				System.out.println(Value);
				
			}
			else if (cellType1.equals(cellType1.NUMERIC)) {
				double Value = (int)cell1.getNumericCellValue();
				int num=(int) Value;
				System.out.println(num);
				
			}
			}
			System.out.println("------Column Data----");
				
				for (int i = 0; i < 5; i++) {
				Row row = sheetAt.getRow(i);
				Cell cell = row.getCell(1);
				CellType cellType = cell.getCellType();
				if (cellType.equals(cellType.STRING)) {
					String Value = cell.getStringCellValue();
					System.out.println(Value);
					
				}
				else if (cellType.equals(cellType.NUMERIC)) {
					double Value = cell.getNumericCellValue();
					int num=(int) Value;
					System.out.println(num);
					
				}
			
			}
			
		}
	
			
		

		public static void main(String[] args) throws Exception {
			
			methodOne();
		}
		
		
	}


