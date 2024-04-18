package apachepoi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class Excell {
	
	@Test
	public void operation() throws IOException {
	
	FileInputStream fis= new FileInputStream("C:\\Users\\HP\\Downloads\\kris.xlsx");
	XSSFWorkbook book =new XSSFWorkbook(fis);
	XSSFSheet sheet=book.getSheet("Krish");
	sheet.createRow(0).createCell(0).setCellValue("helloWorld");
	//inserting of 1 to 10 number
	
	
	for(int i=1;i<=10;i++)
	{
	sheet.createRow(i).createCell(1).setCellValue(i);
	}
	//fetching the values
	for(int i=1;i<=10;i++)
	{
		
		
		System.out.println(sheet.getRow(i).getCell(1).getNumericCellValue());
	}
	
	
	FileOutputStream fos= new FileOutputStream("C:\\Users\\HP\\Downloads\\kris.xlsx");
	book.write(fos);
	
	
}

}
