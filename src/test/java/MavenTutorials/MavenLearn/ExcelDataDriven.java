package MavenTutorials.MavenLearn;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataDriven {

	public static void main(String[] args) throws IOException {

		ArrayList<String> a = new ArrayList<String>();

		// To add values in array list
		// ArrayList<String> a = new ArrayList<String>();
		// Identify file location
		FileInputStream fis = new FileInputStream("C:\\Users\\ani\\Documents\\Selenium\\Resources\\testData.xlsx");
		// Create object of XSSFWorkBook class
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		// identify at which sheet data is present
		int sheetCount = workbook.getNumberOfSheets();

		System.out.println(sheetCount);

		for (int i = 0; i < sheetCount; i++) {
			if (workbook.getSheetAt(i).getSheetName().equalsIgnoreCase("testdata1")) {
				System.out.println("Sheet present");
				// here we will come inside the sheet
				XSSFSheet sheet = workbook.getSheetAt(i);

				// identify column name by using row iterator
				Iterator<Row> row = sheet.iterator();
				Row firstRow = row.next();
				Iterator<Cell> column = firstRow.cellIterator();
				int k = 0;
				int columnfindAt = 0;
				while (column.hasNext()) {
					Cell columnValue = column.next();
					if (columnValue.getStringCellValue().equalsIgnoreCase("TestCases")) {
						columnfindAt = k;

					}
					k++;
				}
				System.out.println("Column find at " + columnfindAt + "th index");

				// after finding column then we will find actual column say password

				while (row.hasNext()) {
					Row r = row.next();
					if (r.getCell(columnfindAt).getStringCellValue().equalsIgnoreCase("Password")) {
						// get the values of specified column name i.e. password
						Iterator<Cell> actualColumn = r.cellIterator();
						while (actualColumn.hasNext()) {
							Cell getValue = actualColumn.next();

							a.add(getValue.getStringCellValue());
						}

						System.out.println(a);

					}
				}

			}
		}

		workbook.close();

	}

}
