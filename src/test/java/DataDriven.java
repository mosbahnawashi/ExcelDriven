import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {

// Identify the testcases cloumn by scanning the entire 1st row
// Once cloumn is identified; then scan testcases cloumn to identify purchase testcase row
// After you grab purchase testcase row; then pull all the data of that row and feed it into test.

	public static void main(String[] args) throws IOException {

	}

	@SuppressWarnings("deprecation")
	public ArrayList<String> getData(String testCaseName) throws IOException {
		ArrayList<String> arraylist = new ArrayList<String>();
		// FileInputStream argument
		FileInputStream fis = new FileInputStream(
				"C:\\Users\\mosba\\OneDrive\\Desktop\\QA\\eclipse workspace\\ExcelDriven\\Book1.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int sheetscount = workbook.getNumberOfSheets();

		for (int i = 0; i < sheetscount; i++) {

			if (workbook.getSheetName(i).equalsIgnoreCase("Testdata")) {
				XSSFSheet sheet = workbook.getSheetAt(i);

				// Identify the testcases cloumn by scanning the entire 1st row
				Iterator<Row> rows = sheet.iterator(); // Sheet is a collection of rows.
				Row firstRow = rows.next();

				Iterator<Cell> ce = firstRow.cellIterator(); // Row is a collection of cells.

				int k = 0;
				int colomn = 0;
				while (ce.hasNext()) {
					Cell value = ce.next();
					if (value.getStringCellValue().equalsIgnoreCase("TestCases")) {
						// Desired column:
						colomn = k;
					}
					k++;
				}
				// System.out.println(colomn);

				// Once cloumn is identified; then scan testcases cloumn to identify purchase
				// testcase row
				while (rows.hasNext()) {
					Row r = rows.next();
					if (r.getCell(colomn).getStringCellValue().equalsIgnoreCase(testCaseName)) {
						// After you grab purchase testcase row; then pull all the data of that row and
						// feed it into test.
						Iterator<Cell> cv = r.cellIterator();
						while (cv.hasNext()) {
							
							Cell c = cv.next();
							//
							
							switch (c.getCellType()) {
							   
							   case NUMERIC:
								   arraylist.add(NumberToTextConverter.toText(c.getNumericCellValue()));
									//arraylist.add(c.getNumericCellValue());
							                 break;
							   case STRING:
								   //String cellvalue = cv.next().getStringCellValue();
									// You can print the results out in the Java console
									// System.out.println(cellvalue);
									// Or you can send it to ArrayList, add the results to the ArrayList, iterate it
									// and print the results.
									//arraylist.add(cellvalue);
									
									c.getStringCellValue();
							                 break;
							default:
								break;
							}
							
							//
							
							/*
							 * if(c.getCellType()==CellType.STRING) { String cellvalue =
							 * cv.next().getStringCellValue(); // You can print the results out in the Java
							 * console // System.out.println(cellvalue); // Or you can send it to
							 * ArrayList,add the results to the ArrayList, iterate it // and print the
							 * results.
							 * 
							 * arraylist.add(cellvalue); } else {
							 * 
							 * arraylist.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							 * //arraylist.add(c.getNumericCellValue());
							 * 
							 * }
							 */
							 
							
						}
					}

				}
			}
		}
		return arraylist;
	}
}