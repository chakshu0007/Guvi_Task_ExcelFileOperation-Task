package task8;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		// Open the work book
		XSSFWorkbook book = new XSSFWorkbook("C:\\Users\\chakshug\\eclipse-workspace\\ExcelFileOperation\\src\\main\\java\\task8\\Employeedetails.xlsx");
		
		//GEt to the sheet
		XSSFSheet sheet = book.getSheet("Sheet1");
		
		//get the no. of rows
		int rowCount = sheet.getLastRowNum();
		
		// get the no. of columns - for that cursor will go to first row and then to last col to count
		int columnCount = sheet.getRow(0).getLastCellNum();
		
		// create 2D array
		String[][] data = new String[rowCount][columnCount];
		
		// get into row
		for (int i=1; i<=rowCount; i++) {
			XSSFRow row = sheet.getRow(i);
			
			// get into columns
			for(int j =0; j<columnCount; j++) {
				XSSFCell cell = row.getCell(j);
				
				// get the value - read the value
			// System.out.println(cell.getStringCellValue()); // only method to read the values - so in Excel do Fixed Width from Data > Text to Columns menu
				
				//to store in a array
				data[i-1][j] =cell.getStringCellValue();
				}
		// System.out.println();
		}
		for (String[] row : data) {
			for (String x : row){
				System.out.println(x +" ");
			}
		}
		book.close();
		
	}

}

