//package comm.zoho;
//
//import org.testng.annotations.Test;
//import org.testng.annotations.BeforeMethod;
//import java.io.FileInputStream;
//import java.io.IOException;
//import java.util.List;
//import java.util.Properties;
//import java.util.concurrent.TimeUnit;
//
//import org.apache.logging.log4j.LogManager;
//import org.apache.logging.log4j.Logger;
//import org.openqa.selenium.By;
//import org.openqa.selenium.WebDriver;
//import org.openqa.selenium.WebElement;
//import org.openqa.selenium.chrome.ChromeDriver;
//import org.openqa.selenium.support.ui.ExpectedConditions;
//import org.openqa.selenium.support.ui.Select;
//import org.openqa.selenium.support.ui.WebDriverWait;
//import org.testng.annotations.AfterTest;
//import org.testng.annotations.BeforeTest;
//import org.testng.annotations.Test;
//
//import io.github.bonigarcia.wdm.WebDriverManager;
//
//public class zohoTest {
//
//	WebDriver driver = null;
//	WebDriverWait wait = null;
//	String path = "./src/test/resources/prints.log";
//	String excelPath="./src/test/resources/zohoData.xlsx";
//	Properties prop;
//	Logger log = LogManager.getLogger(zohoTest.class.getName());
//	Xls_Reader data = new Xls_Reader(excelPath);
//	
//	
//	@BeforeTest
//	public void setUp() throws IOException {
//		WebDriverManager.chromedriver().setup();
//		driver = new ChromeDriver();
//		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
//		driver.manage().window().maximize();
//		
//		// Configuration codes for getting a properties file
//		FileInputStream ip = new FileInputStream(path);
//		prop = new Properties();
//		prop.load(ip);
//		driver.get(prop.getProperty("url"));
//		
//	}
//	
//	@AfterTest
//	public void closeUp() {
//		driver.quit();
//	}
//	
//	@Test
//	public void printExcel() {
//		Select s = new Select(driver.findElement(By.id("recPerPage")));
//		s.selectByIndex(3);
//		
//		wait = new WebDriverWait(driver,5);
//		wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("#reportTab>tbody>:nth-child(20)")));
//		
//		List<WebElement> rows = driver.findElements(By.cssSelector("#reportTab>tbody>tr"));   // -> ozzy's way -> #reportTab>tbody tr
//		List<WebElement> columns = driver.findElements(By.cssSelector("#reportTab>thead>tr>th"));  // -> ozzy's way -> #reportTab>tbody>:nth-child(3) td
//		
//		List<WebElement> id = driver.findElements(By.cssSelector("#reportTab>tbody>tr>:nth-child(1)"));
//		List<WebElement> name = driver.findElements(By.cssSelector("#reportTab>tbody>tr>:nth-child(2)"));
//		List<WebElement> email = driver.findElements(By.cssSelector("#reportTab>tbody>tr>:nth-child(3)"));
//		List<WebElement> phone = driver.findElements(By.cssSelector("#reportTab>tbody>tr>:nth-child(4)"));
//		List<WebElement> salary = driver.findElements(By.cssSelector("#reportTab>tbody>tr>:nth-child(5)"));
//		
//		for(int i=0; i<rows.size(); i++) {
//			data.setCellData("data", "ID", i+2, id.get(i).getText());
//			log.debug("Writing Employee ID-" + (id.get(i).getText()));
//			
//			data.setCellData("data", "NAME", i+2, name.get(i).getText());
//			log.debug("Writing Employee NAME-" + (name.get(i).getText()));
//			
//			data.setCellData("data", "EMAIL", i+2, email.get(i).getText());
//			log.debug("Writing Employee EMAIL-" + (email.get(i).getText()));
//			
//			data.setCellData("data", "PHONE", i+2, phone.get(i).getText());
//			log.debug("Writing Employee PHONE-" + (phone.get(i).getText()));
//			
//			data.setCellData("data", "SALARY", i+2, salary.get(i).getText());
//			log.debug("Writing Employee SALARY-" + (salary.get(i).getText()));
//			
//		}
//		
//	}
//	
//}





package comm.zoho;

import org.testng.annotations.Test;
import org.testng.annotations.AfterTest;


import java.io.FileInputStream;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;


import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import io.github.bonigarcia.wdm.WebDriverManager;


public class zohoTest {
	WebDriver driver;
	
	Logger log = LogManager.getLogger(zohoTest.class.getName());
	WebDriverWait wait = null;
	@AfterTest
	public void after() {
		driver.quit();
	}
	
	@Test
	public void loginTest() throws IOException  {
		
		Properties prop = new Properties();
		FileInputStream ip = new FileInputStream("./src/test/java/comm/zoho/zoho.properties");
		prop.load(ip);
		Xls_Reader xlBook = new Xls_Reader("./src/test/resources/zohoData.xlsx");
		
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.get(prop.getProperty("url"));
		log.debug("Opening the browser");
		driver.manage().window().maximize();
		Select sel = new Select(driver.findElement(By.id("recPerPage")));
		sel.selectByIndex(3);
		wait=new WebDriverWait(driver, 5);
	    wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("#reportTab>tbody>:nth-child(50)")));
		log.info("Selecting dropdown");
		
		List<WebElement> allHeaders = driver.findElements(By.cssSelector("#reportTab>thead>tr>th"));
		log.info("Getting text from headers");
		List<WebElement> firstCol = driver.findElements(By.xpath("//*[@class='reportTable']/tbody/tr/td[1]"));
		List<WebElement> secondCol = driver.findElements(By.xpath("//*[@class='reportTable']/tbody/tr/td[2]"));
		List<WebElement> thirdCol = driver.findElements(By.xpath("//*[@class='reportTable']/tbody/tr/td[3]"));
		List<WebElement> fourthCol = driver.findElements(By.xpath("//*[@class='reportTable']/tbody/tr/td[4]"));
		List<WebElement> fifthCol = driver.findElements(By.xpath("//*[@class='reportTable']/tbody/tr/td[5]"));
		
		List <String> headers = new ArrayList <String>();
		for (int i=0; i<allHeaders.size(); i++) {
			headers.add(allHeaders.get(i).getText());
			String each1 = headers.get(i);
			xlBook.addColumn("data1", each1);
			log.info("Writing Header "+each1);
			for(int j=0; j<100; j++) {
				if(i==0) {
					String result = firstCol.get(j).getText();
					xlBook.setCellData("data1", each1, j+2, result );
					log.debug("Writing CellData "+result+" in "+each1);
				}else if (i==1) {
					String result = secondCol.get(j).getText();
					xlBook.setCellData("data1", each1, j+2, result );
					log.debug("Writing CellData "+result+" in "+each1);
				}else if (i==2) {
					String result = thirdCol.get(j).getText();
					xlBook.setCellData("data1", each1, j+2, result );
					log.debug("Writing CellData "+result+" in "+each1);
				}else if (i==3) {
					String result = fourthCol.get(j).getText();
					xlBook.setCellData("data1", each1, j+2, result );
					log.debug("Writing CellData "+result+" in "+each1);
				}else if (i==4) {
					String result = fifthCol.get(j).getText();
					xlBook.setCellData("data1", each1, j+2, result );
					log.info("Writing CellData "+result+" in "+each1);
				}
			
				
			}
		}
		
	}

}

