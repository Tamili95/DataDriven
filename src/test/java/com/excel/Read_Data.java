package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_Data {

	public static void particularCellData() throws Throwable {

		File f = new File("C:\\Users\\Tamil\\eclipse-workspace\\DataDriven\\UserDetails.xlsx");
		FileInputStream fis = new FileInputStream(f);

		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheetAt = wb.getSheetAt(0);
		Row row = sheetAt.getRow(0);
		Cell cell = row.getCell(0);

		CellType cellType = cell.getCellType();
		if (cellType.equals(CellType.STRING)) {

			String stringCellValue = cell.getStringCellValue();
			System.out.println(stringCellValue);

		} else if (cellType.equals(CellType.NUMERIC)) {

			double numericCellValue = cell.getNumericCellValue();
			int value = (int) numericCellValue;
			System.out.println(value);

		}

	}

	public static void allData() throws Throwable {

		File f = new File("C:\\Users\\Tamil\\eclipse-workspace\\DataDriven\\UserDetails.xlsx");
		FileInputStream fis = new FileInputStream(f);

		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheetAt = wb.getSheetAt(0);
		for (int i = 0; i < sheetAt.getPhysicalNumberOfRows(); i++) {

			Row row = sheetAt.getRow(i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {

				Cell cell = row.getCell(j);
				CellType cellType = cell.getCellType();

				if (cellType.equals(CellType.STRING)) {

					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue);

				} else if (cellType.equals(CellType.NUMERIC)) {

					double numericCellValue = cell.getNumericCellValue();
					int value = (int) numericCellValue;
					System.out.println(value);

				}

			}
		}

	}

	public static void particularRowData(int rows) throws Throwable {

		File f = new File("C:\\Users\\Tamil\\eclipse-workspace\\DataDriven\\UserDetails.xlsx");
		FileInputStream fis = new FileInputStream(f);

		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheetAt = wb.getSheetAt(0);

		Row row = sheetAt.getRow(rows);
		for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {

			Cell cell = row.getCell(j);
			CellType cellType = cell.getCellType();
			if (cellType.equals(CellType.STRING)) {

				String stringCellValue = cell.getStringCellValue();
				System.out.println(stringCellValue);

			} else if (cellType.equals(CellType.NUMERIC)) {

				double numericCellValue = cell.getNumericCellValue();
				int value = (int) numericCellValue;
				System.out.println(value);
			}

		}

	}
	
	
	public static void particularColumnData(int column) throws Throwable {

		File f = new File("C:\\Users\\Tamil\\eclipse-workspace\\DataDriven\\UserDetails.xlsx");
		FileInputStream fis = new FileInputStream(f);

		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheetAt = wb.getSheetAt(0);

		for (int i = 0; i < sheetAt.getPhysicalNumberOfRows(); i++) {
			
		Row row = sheetAt.getRow(i);
	
			Cell cell = row.getCell(column);
			CellType cellType = cell.getCellType();
			if (cellType.equals(CellType.STRING)) {

				String stringCellValue = cell.getStringCellValue();
				System.out.println(stringCellValue);

			} else if (cellType.equals(CellType.NUMERIC)) {

				double numericCellValue = cell.getNumericCellValue();
				int value = (int) numericCellValue;
				System.out.println(value);
			}
		}
		

	}

	public static void main(String[] args) throws Throwable {

//		particularCellData();
//		allData();
		particularRowData(2);
		particularColumnData(1);

	}

}
