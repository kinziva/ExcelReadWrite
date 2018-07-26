package tests;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelWrite {

	public static void main(String[] args) throws Exception {

		String testfilePath = "./src/test/resources/testData/Excel.xlsx";
		FileInputStream inStream = new FileInputStream(testfilePath);

		Workbook workbook = WorkbookFactory.create(inStream);
		Sheet sheet = workbook.getSheetAt(0);

		Row row = sheet.getRow(0);
		int rowCount = sheet.getPhysicalNumberOfRows();

		int cellNumber = row.getPhysicalNumberOfCells();
		System.out.println("Col number " + cellNumber);
		
		for (int rowNum = 1; rowNum < rowCount ; rowNum++) {

			System.out.println(rowNum + " - " + sheet.getRow(rowNum).getCell(3));
			Cell loopjob = sheet.getRow(rowNum).getCell(3);
			loopjob.setCellValue("SDET");
		}

		Cell job = sheet.getRow(1).getCell(0);
		job.setCellValue("Asli");
		FileOutputStream outStream = new FileOutputStream(testfilePath);

		workbook.write(outStream);
		outStream.close();
		workbook.close();
		inStream.close();
	}
}
