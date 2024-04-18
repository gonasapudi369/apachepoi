package apachepoi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.Duration;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class Hrmlogin {
	WebDriver driver;
	FileInputStream fis;
	XSSFWorkbook book;
	XSSFSheet  sheet;
	FileOutputStream fos;
	@BeforeClass
	public void LaunchApp() {
		
		 driver= (WebDriver) new ChromeDriver();
		 driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login");
		 driver.manage().window().maximize();
		 driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
		 
		
		
	}
	@Test
	public void operation() 
	{
		fis=new FileInputStream("C:\\Users\\HP\\OneDrive\\Desktop\\vamsi.xlsx");
		book=new XSSFWorkbook(fis);
		sheet=book.getSheet("krishna");
		
driver.findElement(By.name("username")).sendKeys(sheet.getRow(1).getCell(0).getStringCellValue());
driver.findElement(By.name("password")).sendKeys(sheet.getRow(1).getCell(1).getStringCellValue());
driver.findElement(By.xpath("//button[normalize-space()='Login']")).click();
   Thread.sleep(3000);
   
      String expectedUrl=sheet.getRow(1).getCell(2).getStringCellValue();
   String actualUrl=driver.getCurrentUrl();
   sheet.getRow(1).createCell(3).setCellValue(actualUrl);
   
   if(expectedUrl.contains(actualUrl)) 
   {
	   System.out.println("pass");
	   sheet.getRow(0).getCell(3).setCellValue("pass");
	   
   }
   else {
	   System.out.println("Fail");
	   sheet.getRow(0).getCell(3).setCellValue("Fail");
	   
  
   }
   fos=new FileOutputStream("C:\\Users\\HP\\OneDrive\\Desktop\\vamsi.xlsx");
   book.write(fos);
	}
	
	

	
	
}
