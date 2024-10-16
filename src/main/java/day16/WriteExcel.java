package day16;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {
	public static void main(String[] args) throws IOException {
		
		// First create a work book
		XSSFWorkbook book = new XSSFWorkbook();
		
		// Create the sheet inside book
		XSSFSheet sheet = book.createSheet("details");
		
		// Store the details -> Name , Age , City 
		// using Object because in ArrayList this Object store all tyes of datatype
		// [][] rows and columns
		Object[][] data = {
				{"Name","Age","City"},
				{"Aaaa",20,"Delhi"},
				{"Bbbb",25,"Chennai"},
				{"Cccc",30,"Mumbai"}
		};
		
		// initially set row count to 0, everytime we enter data in 1 row rowCount is going to increase
		int rowCount = 0;
		
		// using for each loop to get into each row
		// focusing on 1 row will give 1-d array 
		for(Object[] row1 : data) {  		 // this means for each row in data
			// now get into each cell
			XSSFRow row = sheet.createRow(rowCount++);
			
			int columnCount = 0;
			
			for(Object col : row1) {		// this means for each column in a row
			
			XSSFCell cell =	row.createCell(columnCount++);
			
			// checking the type of data and making the entry
			if(col instanceof String) {
				cell.setCellValue((String)col); // here converting the col from Object to String
				
			}else if(col instanceof Integer) {
				cell.setCellValue((Integer)col); // because only 2 dataypes are present,else need to convert all
			}
			}
		}
		// fileoutputstream to write in xlsx with path given '\\'
	try {
		FileOutputStream output = new FileOutputStream("C:\\Users\\chakshug\\eclipse-workspace\\ExcelFileOperation\\src\\main\\java\\day16\\Studentdetails.xlsx");
			// If file found then write in the book
		book.write(output);
	} catch (Exception e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	
	}
	// IMP to CLOSE THE OPENED BOOK
	book.close();
	
	}

}
