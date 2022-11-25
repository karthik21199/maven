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

public class TestingRowColumn {
	
	
	public static void checkingData() throws IOException {
	
		File f = new File("C:\\Users\\LENOVO\\Desktop\\Project_11Am\\Book1.xlsx");
		
		FileInputStream first = new FileInputStream(f);
		
		Workbook b = new XSSFWorkbook(first);
		Sheet sheetAt = b.getSheetAt(0);
		for (int i = 0; i < 10; i++) {
			Row row = sheetAt.getRow(i);
			for (int j = 0; j < 2; j++) {
				j=++j;
				Cell cell = row.getCell(j);
				CellType cellType = cell.getCellType();
				if (cellType.equals(cellType.STRING)) {
					String Value = cell.getStringCellValue();
					System.out.println(Value);
				}
				else if (cellType.equals(cellType.NUMERIC)) {
					double Value = cell.getNumericCellValue();
					int num = (int) Value;
					System.out.println(num);
					
				}	
				}
			}
		
		
		
		
		
	}
	
	public static void main(String[] args) throws Exception {
	
		checkingData();
	}

}
