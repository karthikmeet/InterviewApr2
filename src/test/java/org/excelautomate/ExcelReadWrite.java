package org.excelautomate;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Ignore;
import org.testng.annotations.Test;

public class ExcelReadWrite {

	@Ignore
	@Test
	public void excelRead() throws IOException {
		// xls ms 97 (HSSF)
		// xlsx 2003 to till date (XSSF)

		System.out.println("Reading my excel");

		File f = new File(System.getProperty("user.dir") + "/target/Student Details - Project Apr2 batch.xlsx");
		FileInputStream input = new FileInputStream(f);
		XSSFWorkbook workbook = new XSSFWorkbook(input);
		XSSFSheet sheet = workbook.getSheet("Project Apr2 batch");
		int totalRows = sheet.getPhysicalNumberOfRows();

		for (int i = 0; i < totalRows; i++) {
			XSSFRow row = sheet.getRow(i);

			int physicalnumberofcells = row.getPhysicalNumberOfCells();
			for (int j = 0; j < physicalnumberofcells; j++) {
				XSSFCell cell = row.getCell(j);

				if (cell.getCellType() == CellType.NUMERIC) {
					double numericCellValue = cell.getNumericCellValue();
					System.out.println(numericCellValue + " ");
				} else {
					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue + " ");
				}
			}
			System.out.println(" ");
		}
		workbook.close();
	}
	
	@Test
	public void writeExcel() throws IOException {
		File f = new File(System.getProperty("user.dir") + "/target/Student Details - Project Apr2 batch.xlsx");
		FileInputStream input = new FileInputStream(f);
		XSSFWorkbook workbook = new XSSFWorkbook(input);
		XSSFSheet sheet = workbook.getSheet("Project Apr2 batch");
		// int totalRows = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		row.getCell(0).setCellValue("Update cell value");
		row.createCell(10);
		row.getCell(10).setCellValue("Write new cell value");		
		FileOutputStream output = new FileOutputStream(f);
		workbook.write(output);
		workbook.close();
		output.close();
	}
}
