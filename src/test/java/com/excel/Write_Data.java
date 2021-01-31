package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write_Data {

	public static void createSheet() throws Throwable {

		File f = new File("C:\\Users\\Tamil\\Desktop\\UserDetails.xlsx");
		FileInputStream fis = new FileInputStream(f);

		Workbook wb = new XSSFWorkbook(fis);
		Sheet createSheet = wb.createSheet("Amazon Users");

		Row createRow = createSheet.createRow(0);
		Cell createCell = createRow.createCell(0);

		createCell.setCellValue("User Name");

		wb.getSheet("Amazon Users").getRow(0).createCell(1).setCellValue("Password");
		wb.getSheet("Amazon Users").createRow(1).createCell(0).setCellValue("Jamie");
		wb.getSheet("Amazon Users").getRow(1).createCell(1).setCellValue("484849");

		FileOutputStream fos = new FileOutputStream(f);

		wb.write(fos);
		wb.close();

	}

	public static void main(String[] args) throws Throwable {

		createSheet();
	}

}
