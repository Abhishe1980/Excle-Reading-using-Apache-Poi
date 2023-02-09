package exceloperation;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.*;

public class Readingexcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		String excelFilePath = ".\\Datafile\\Countries.xlsx";

		FileInputStream inputStream = new FileInputStream(excelFilePath);

		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

		// XSSFSheet sheet= workbook.getSheet("sheet1");
		XSSFSheet sheet = workbook.getSheetAt(0);

		// Using For Loop
		/*
		 * int rows = sheet.getLastRowNum(); int cols=sheet.getRow(1).getLastCellNum();
		 * 
		 * for(int r=0;r<=rows;r++) { XSSFRow row=sheet.getRow(r);
		 * 
		 * for(int c=0;c<=cols;c++) {
		 * 
		 * XSSFCell cell=row.getCell(c); switch(cell.getCellType()) { case
		 * STRING:System.out.print(cell.getStringCellValue()); break; case
		 * NUMERIC:System.out.print(cell.getNumericCellValue()); break; case
		 * BOOLEAN:System.out.print(cell.getBooleanCellValue()); break; }
		 * System.out.println(" | "); } System.out.println(); }
		 */

		// Itrator method
		Iterator iterator = sheet.iterator();
		while (iterator.hasNext()) {
			XSSFRow row = (XSSFRow) iterator.next();
			Iterator cellIterator = row.cellIterator();

			while (cellIterator.hasNext()) {

				XSSFCell cell = (XSSFCell) cellIterator.next();

				switch (cell.getCellType()) {
				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				}
				System.out.print("   |   ");

			}
			
			System.out.println();

		}

	}

}
