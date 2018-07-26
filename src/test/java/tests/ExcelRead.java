package tests;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelRead {
	
	public static void main(String[] args) throws Exception {
		
		String filePath = "./src/test/resources/testData/Excel.xlsx";
		FileInputStream inStream = new FileInputStream(filePath);
		
		Workbook workbook = WorkbookFactory.create(inStream);
		Sheet sheet1 = workbook.getSheetAt(0);
		Sheet sheet2 = workbook.getSheet("Sheet1");
		Row row =sheet1.getRow(0);
		Row row2 =sheet2.getRow(0);
		Cell cell2 = row2.getCell(0);
		Cell cell = row.getCell(0);
		System.out.println(cell.toString());
		System.out.println(cell2.toString());
		
		int rowCount = sheet1.getLastRowNum();
		System.out.println(rowCount);
		
		int rowCountp = sheet1.getPhysicalNumberOfRows();
		System.out.println(rowCountp);
		
		for (int rowNum= 1; rowNum < rowCountp; rowNum++) {
			
			System.out.println(rowNum + " - " + sheet1.getRow(rowNum).getCell(3));
		}
		
		workbook.close();
		inStream.close();
	}
}
