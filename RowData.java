package com.DataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class RowData {

	public static void particularRow() throws Exception {
		File f = new File("C:\\\\Users\\\\LENOVO\\\\Desktop\\\\Project_11Am\\\\Book2.xlsx");
		FileInputStream web = new FileInputStream(f);
		Workbook t = new XSSFWorkbook(web);
		Sheet sheetAt = t.getSheetAt(0);
		for (int i = 1; i < 2; i++) {
			Row row = sheetAt.getRow(i);
			for (int j = 0; j <1; j++) {
				//j=++j;
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
			particularRow();
			
		}
		
	}
