package qsp_tutorial;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Write_Data_Excel_DDT {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try {
			FileInputStream fis = new FileInputStream("C://Users//Naveen//Documents//Excel_Data_for_Selenium//Write_Data_DDT.xlsx");
			Workbook book = WorkbookFactory.create(fis);
			Sheet sh = book.getSheet("Sheet1");
			Row row = sh.createRow(0);
			Cell cell = row.createCell(0);
			cell.setCellValue("NAVEENA");
			FileOutputStream fos = new FileOutputStream("C://Users//Naveen//Documents//Excel_Data_for_Selenium//Write_Data_DDT.xlsx");
			book.write(fos);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

}
