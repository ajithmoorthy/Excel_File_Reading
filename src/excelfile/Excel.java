package excelfile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.Iterator;
import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**Creating the class for the Implement the Excel file Reading */
public class Excel {
public static void main(String[] args) throws BiffException, IOException {

	String data;
	//File class is  
	File file=new File("C:\\Users\\ajith.periyasamy\\Downloads\\Book.xls.xls");
	//Excel file can access by the using the workbook class 
	 Workbook workbook = Workbook.getWorkbook(file);
	 //sheet using to get the sheets of the Book.xlsx file 
	Sheet sheet=workbook.getSheet(0);
	//get the row count and column count
	int colCount=sheet.getColumns();
	int rowCount=sheet.getRows();
	//for loop for the access the Excel file using the row and column index
	for(int initial=0; initial<rowCount; initial++)
	{
		for(int Count=0; Count<colCount; Count++)
		{
			//Get the content of the sheet using getcell() and getcontent() methods 
			data=sheet.getCell(Count,initial).getContents(); 
			System.out.print(data+"\t\t");
		}
		System.out.println();
	}
	workbook.close();
	
	}
}
