package Execute;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Proxy;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class Sim_google 
{
	WebDriver driver;
	File src;
	File file;
	FileInputStream fis;
	XSSFWorkbook workbook;
	XSSFSheet shet;
	XSSFWorkbook wb;
	XSSFSheet sheet;
	XSSFRow row;
	FileOutputStream fos;
	String name;
	String address;
	String number;
	String website;
	AtomicInteger count = new AtomicInteger(1);
	@BeforeTest
	public void g() throws InterruptedException, IOException
	{
		Proxy p=new Proxy();
	    p.setHttpProxy("45.77.156.187:8080");
	    DesiredCapabilities cap=new DesiredCapabilities();     
	      // Pass proxy object p
	    cap.setCapability(CapabilityType.PROXY, p);
	   System.setProperty("webdriver.chrome.driver", "D:\\Software\\Selenium\\New Driver\\Chrome Driver 83\\chromedriver.exe");
	    driver = new ChromeDriver();
		//FirefoxOptions foptions = new FirefoxOptions();
	  // foptions.setCapability("marionette", false);
	   // driver = new FirefoxDriver(foptions);
	    driver.manage().window().maximize(); 
	    src = new File("D:\\Software\\Database_2020\\db shop database\\All India Medical june 2019\\chemistnashik.xlsx");
		wb = new XSSFWorkbook(); 
		sheet = wb.createSheet("res_details");	
		row = sheet.createRow((short)0);  
		row.createCell(0).setCellValue("Name");
		row.createCell(1).setCellValue("Address");
		row.createCell(2).setCellValue("Number");
		row.createCell(3).setCellValue("website");	
	    fos = new FileOutputStream(src); 
		wb.write(fos);
		driver.navigate().to("https://www.google.com/search?tbm=lcl&q=chemists+in+nashik&spell=1&sa=X&ved=0ahUKEwiawMWjqfrpAhUS8XMBHSLBCqkQBQgdKAA&biw=1536&bih=710&dpr=1.25#rlfi=hd:;si:;mv:[[20.0186682,73.84549799999999],[19.9245658,73.7289721]];tbs:lrf:!1m4!1u3!2m2!3m1!1e1!2m1!1e3!3sIAE,lf:1,lf_ui:3");
		Thread.sleep(3000);
	}
	@Test
	public void h() throws InterruptedException, IOException
	{	
	List<WebElement> list = driver.findElements(By.cssSelector(".dbg0pd>div"));
	for(int i =0; i<list.size(); i++)
	{
		((JavascriptExecutor)driver).executeScript("scroll(10,400)");
		list.get(i).click();
		Thread.sleep(3000);
		row = sheet.createRow((short)count.getAndIncrement());
		
		fos = new FileOutputStream(src); 
		wb.write(fos);
		try
		{
			name = list.get(i).getText();
			row.createCell(0).setCellValue(name);
			fos = new FileOutputStream(src); 
			wb.write(fos);
		}
		catch(Exception e2)
		{
			try
			{
				 name = driver.findElement(By.xpath(".//*[@id='akp_uid_2']/div/div/div/div/div/div[1]/div/div[1]/div[2]/div[1]/div/div[1]/div/div[1]/div/span")).getText();
				row.createCell(0).setCellValue(name);
				fos = new FileOutputStream(src); 
				wb.write(fos);
			}
			catch(Exception e3)
			{
				try
				{
				 name = driver.findElement(By.cssSelector(".FxvUNb.kno-ecr-pt.kno-fb-ctx.PPT5v.hNKfZe>span")).getText();
					row.createCell(0).setCellValue(name);
					fos = new FileOutputStream(src); 
					wb.write(fos);
				}
				catch(Exception e8)
				{
					
				}
			}
		}
		try
		{
			address = driver.findElement(By.xpath("//span[@class='LrzXr']")).getText();
			row.createCell(1).setCellValue(address);
			fos = new FileOutputStream(src); 
			wb.write(fos);
			
		}
		catch(Exception e4)
		{
			try
			{
			 address = driver.findElement(By.cssSelector(".LrzXr")).getText();
				row.createCell(1).setCellValue(address);
				fos = new FileOutputStream(src); 
				wb.write(fos);
			}
			catch(Exception e5)
			{
				try
				{
					address = driver.findElement(By.xpath(".//*[@id='akp_uid_2']/div/div/div/div/div/div[1]/div/div[1]/div[2]/div[2]/div/div[2]/div/div/span[2]")).getText();
					row.createCell(1).setCellValue(address);
					fos = new FileOutputStream(src); 
					wb.write(fos);
				}
				catch(Exception ee)
				{
					
				}
			}
		}
		try
		{
			String number = driver.findElement(By.cssSelector(".LrzXr.zdqRlf.kno-fv")).getText();
			row.createCell(2).setCellValue(number);
			fos = new FileOutputStream(src); 
			wb.write(fos);
		}
		catch(Exception e6)
		{
			try
			{
				String number = driver.findElement(By.cssSelector(".LrzXr.zdqRlf.kno-fv>span>span")).getText();
				row.createCell(2).setCellValue(number);
				fos = new FileOutputStream(src); 
				wb.write(fos);
			}
			catch(Exception e7)
			{
				try
				{
					String number = driver.findElement(By.xpath("//*[@id='akp_tsuid5']/div/div[1]/div/div/div/div[1]/div/div[1]/div/div[5]/div/div/span[2]/a/span")).getText();
					row.createCell(2).setCellValue(number);
					fos = new FileOutputStream(src); 
					wb.write(fos);	
				}
				catch(Exception e8)
				{
					
				}
			}
		}
		
		try
		{
		 website = driver.findElement(By.xpath(".//*[@id='akp_uid_2']/div/div/div/div/div/div[1]/div/div[1]/div[2]/div[1]/div/div[1]/div/div[2]/div[1]/a")).getAttribute("href");
		row.createCell(3).setCellValue(website);
		fos = new FileOutputStream(src); 
		wb.write(fos);
		}
		catch(Exception e)
		{
			try
			{
				 website = driver.findElement(By.cssSelector(".CL9Uqc.ab_button")).getAttribute("href");
				row.createCell(3).setCellValue(website);
				fos = new FileOutputStream(src); 
				wb.write(fos);
			}
			catch(Exception e9)
			{
				try
				{
					website = driver.findElement(By.xpath("//a[contains(text(),'Website')]")).getAttribute("href");
					row.createCell(3).setCellValue(website);
					fos = new FileOutputStream(src); 
					wb.write(fos);
				}
				catch(Exception e10)
				{
					
				}
			}
		}
		
	}
	  
	list.clear();
	 k();
	}
public void k() throws InterruptedException, IOException
{
	driver.findElement(By.cssSelector("#pnnext>span")).click();
	   Thread.sleep(3000);
	   h();
}
}