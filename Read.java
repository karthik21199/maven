package com.DataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read {
	
	public static void readAllData() throws IOException{
		
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
		       
			
			
			
		}
		public static void main(String[] args) throws IOException {
			
			readAllData();
		}
		
		
	}

	


