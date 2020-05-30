package Qedge;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Reporter;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class AppTest {
WebDriver driver;
Workbook wb;
Sheet ws;
Row row;
FileInputStream fi;
FileOutputStream fo;
@BeforeTest
public void setUp()
{
driver=new ChromeDriver();
}
@Test
public void login()throws Throwable
{
	fi=new FileInputStream("g://logindata.xlsx");
	wb=WorkbookFactory.create(fi);
	ws=wb.getSheet("Login");
	int rc=ws.getLastRowNum();
	for(int i=1;i<=rc;i++)
	{
		driver.get("http://orangehrm.qedgetech.com/");
		driver.manage().window().maximize();
		String username=ws.getRow(i).getCell(0).getStringCellValue();
		String password=ws.getRow(i).getCell(1).getStringCellValue();
		driver.findElement(By.name("txtUsername")).sendKeys(username);
		driver.findElement(By.name("txtPassword")).sendKeys(password);
		driver.findElement(By.name("Submit")).click();
		Thread.sleep(5000);
		if(driver.getCurrentUrl().contains("dash"))
		{
			Reporter.log("Login Success",true);
			ws.getRow(i).createCell(2).setCellValue("Login Success");
		}
		else
		{
			Reporter.log("Login Fail",true);
			ws.getRow(i).createCell(2).setCellValue("Login Fail");	
		}
	}
	fi.close();
	fo=new FileOutputStream("G://Results.xlsx");
	wb.write(fo);
	fo.close();
	wb.close();
}
@AfterTest
public void tearDown()throws Throwable
{
	Thread.sleep(5000);
	driver.close();
}
}













