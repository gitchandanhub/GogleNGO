
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
import org.openqa.selenium.Keys;
import org.openqa.selenium.Proxy;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class SchoolMoreComplex_google 
{
	String website;
	String City = "";
	String category = "";
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
	String Founded = "";
	String Affiliation = "";
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
	    src = new File("D:\\Software\\Database_2020\\db shop database\\allindiaschooldatabase (1)\\May 2020\\Google\\juneschool2020_2.xlsx");
		wb = new XSSFWorkbook(); 
		sheet = wb.createSheet("ME_details");	
		row = sheet.createRow((short)0);  
		row.createCell(0).setCellValue("Name");
		row.createCell(1).setCellValue("Address");
		row.createCell(2).setCellValue("Number");
		row.createCell(3).setCellValue("website");
		row.createCell(4).setCellValue("City");
		row.createCell(5).setCellValue("Category");
		row.createCell(6).setCellValue("Founded");
		row.createCell(7).setCellValue("Affiliation");
	    fos = new FileOutputStream(src); 
		wb.write(fos);
		driver.navigate().to("https://www.google.com/search?sz=0&tbm=lcl&ei=sabjXtn8FIX39QPLzIvQAQ&q=Schools%20in%20Haryana&oq=Schools+in+Haryana&gs_l=psy-ab.3...0.0.0.43069.0.0.0.0.0.0.0.0..0.0....0...1c..64.psy-ab..0.0.0....0.O5qAoCJK72c&tbs=lrf:!1m4!1u3!2m2!3m1!1e1!1m4!1u2!2m2!2m1!1e1!2m1!1e2!2m1!1e3!2m4!1e17!4m2!17m1!1e2!3sIAE,lf:1,lf_ui:2&rlst=f#rlfi=hd:;si:;mv:[[29.3367429,77.4264778],[28.1636219,76.8537856]];tbs:lrf:!1m4!1u3!2m2!3m1!1e1!1m4!1u2!2m2!2m1!1e1!2m1!1e2!2m1!1e3!2m4!1e17!4m2!17m1!1e2!3sIAE,lf:1,lf_ui:2");
		Thread.sleep(180000);
	}
	@Test
	public void a() throws IOException, InterruptedException
	{
		for(int i =86; i<22009; i++)
		{
			City = Readwriteexcel.readexcel("D:\\Software\\Database_2020\\db shop database\\allindiaschooldatabase (1)\\May 2020\\Google\\schoolurl.xlsx", 0, i, 0);
			String url = Readwriteexcel.readexcel("D:\\Software\\Database_2020\\db shop database\\allindiaschooldatabase (1)\\May 2020\\Google\\schoolurl.xlsx", 0, i, 1);
			category = Readwriteexcel.readexcel("D:\\Software\\Database_2020\\db shop database\\allindiaschooldatabase (1)\\May 2020\\Google\\schoolurl.xlsx", 0, i, 2);
			driver.navigate().to(url);
			Thread.sleep(0000);
			System.out.println("Page no : -" + i);
			Thread.sleep(3000);
			h();
		}
	}
	public void h() throws InterruptedException, IOException
	{	
		
	List<WebElement> list = driver.findElements(By.cssSelector(".dbg0pd>div"));
	if(list.size()==0)
	{
		driver.findElement(By.xpath("//*[@id=\"lst-ib\"]")).click();
		Actions act = new Actions(driver); 
		act.sendKeys(Keys.ENTER).build().perform(); 
		m();
		
	}
	else
	{
		m();
	}
	list.clear();
	}
	public void m() throws InterruptedException, IOException
	{
		List<WebElement> list2 = driver.findElements(By.cssSelector(".dbg0pd>div"));
	
	for(int i =0; i<list2.size(); i++)
	{
		((JavascriptExecutor)driver).executeScript("scroll(10,400)");
		list2.get(i).click();
		Thread.sleep(3000);
		row = sheet.createRow((short)count.getAndIncrement());
		row.createCell(4).setCellValue(City);
		row.createCell(5).setCellValue(category);
		fos = new FileOutputStream(src); 
		wb.write(fos);
		try
		{
			String name = list2.get(i).getText();
			row.createCell(0).setCellValue(name);
			fos = new FileOutputStream(src); 
			wb.write(fos);
		}
		catch(Exception e2)
		{
			try
			{
				String name = driver.findElement(By.xpath(".//*[@id='akp_uid_2']/div/div/div/div/div/div[1]/div/div[1]/div[2]/div[1]/div/div[1]/div/div[1]/div/span")).getText();
				row.createCell(0).setCellValue(name);
				fos = new FileOutputStream(src); 
				wb.write(fos);
			}
			catch(Exception e3)
			{
				try
				{
					String name = driver.findElement(By.cssSelector(".FxvUNb.kno-ecr-pt.kno-fb-ctx.PPT5v.hNKfZe>span")).getText();
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
			String address = driver.findElement(By.xpath(".//*[@id='akp_uid_2']/div/div/div/div/div/div[1]/div/div[1]/div[2]/div[2]/div/div[2]/div/div/span[2]")).getText();
			row.createCell(1).setCellValue(address);
			fos = new FileOutputStream(src); 
			wb.write(fos);
		}
		catch(Exception e4)
		{
			try
			{
				String address = driver.findElement(By.cssSelector(".LrzXr")).getText();
				row.createCell(1).setCellValue(address);
				fos = new FileOutputStream(src); 
				wb.write(fos);
			}
			catch(Exception e5)
			{
				
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
				String number = driver.findElement(By.xpath("//span[@class='zgWrF']")).getText();
				row.createCell(2).setCellValue(number);
				fos = new FileOutputStream(src); 
				wb.write(fos);
			}
			catch(Exception e7)
			{
				try
				{
					String number = driver.findElement(By.cssSelector(".LrzXr.zdqRlf.kno-fv>span>span")).getText();
					row.createCell(2).setCellValue(number);
					fos = new FileOutputStream(src); 
					wb.write(fos);	
				}
				catch(Exception e8)
				{
					try
					{
						String number = driver.findElement(By.xpath(".//*[@id='akp_uid_2']/div/div/div/div/div/div[1]/div/div[1]/div[2]/div[2]/div/div[4]/div/div/span[2]/span/span")).getText();
						row.createCell(2).setCellValue(number);
						fos = new FileOutputStream(src); 
						wb.write(fos);
					}
					catch(Exception e11)
					{
						
					}
				}
			}
		}
		try
		{
			website = driver.findElement(By.xpath("//a[contains(text(),'Website')]")).getAttribute("href");
			row.createCell(3).setCellValue(website);
			fos = new FileOutputStream(src); 
			wb.write(fos);
		}
		catch(Exception e)
		{
			try
			{
				website = driver.findElement(By.xpath(".//*[@id='akp_uid_2']/div/div/div/div/div/div[1]/div/div[1]/div[2]/div[1]/div/div[1]/div/div[2]/div[1]/a")).getAttribute("href");
				row.createCell(3).setCellValue(website);
				fos = new FileOutputStream(src); 
				wb.write(fos);
			}
			catch(Exception e9)
			{
				try
				{
					website = driver.findElement(By.cssSelector(".CL9Uqc.ab_button")).getAttribute("href");
					row.createCell(3).setCellValue(website);
					fos = new FileOutputStream(src); 
					wb.write(fos);
				}
				catch(Exception e10)
				{
					
				}
			}
		}
		
			
		try
		{
		Founded = driver.findElement(By.xpath("/html[1]/body[1]/div[6]/div[3]/div[9]/div[1]/div[3]/div[1]/div[1]/div[1]/div[2]/async-local-kp[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[4]/div[1]/div[1]/span[2]")).getText();
		row.createCell(6).setCellValue(Founded);
		fos = new FileOutputStream(src); 
		wb.write(fos);
		}
		catch(Exception e)
		{
			try
			{
				Founded = driver.findElement(By.cssSelector("body.srp.tbo.vasq:nth-child(2) div.rfli.mdm:nth-child(3) div.BBj6N:nth-child(12) div.col:nth-child(3) div.rhstc5 div.r-iYSmFJGakXPQ:nth-child(2) div.iYSmFJGakXPQ-jWD6pBvCV60.h2yBfgNjGpc__inline-item-view.r-iCW_xhUARSoI async-local-kp.iCW_xhUARSoI-AjwZTeXvKls.r-iMpgniV_zCxk div.iMpgniV_zCxk-AjwZTeXvKls.y.yf div.immersive-container:nth-child(2) div.akp_tsuid3.akp-el div.t3HED div.g.kno-kp.mnr-c.g-blk div.kp-blk.knowledge-panel.Wnoohf.OJXvsb div.xpdopen div.ifM9O:nth-child(1) div:nth-child(2) div.mod:nth-child(4) div.Z1hOCe div.zloOqf.PZPZlf.kno-fb-ctx > span.LrzXr.kno-fv")).getText();
				row.createCell(6).setCellValue(Founded);
				fos = new FileOutputStream(src); 
				wb.write(fos);
			}
			catch(Exception e9)
			{
				
			}
		}
		try
		{
	Affiliation = driver.findElement(By.xpath("//*[@id='akp_tsuid3']/div/div/div/div/div/div[1]/div/div[1]/div[2]/div[5]/div/div/span[2]")).getText();
		row.createCell(7).setCellValue(Affiliation);
		fos = new FileOutputStream(src); 
		wb.write(fos);
		}
		catch(Exception e)
		{
			try
			{
				Affiliation = driver.findElement(By.cssSelector("body.srp.tbo.vasq:nth-child(2) div.rfli.mdm:nth-child(3) div.BBj6N:nth-child(12) div.col:nth-child(3) div.rhstc5 div.r-iYSmFJGakXPQ:nth-child(2) div.iYSmFJGakXPQ-jWD6pBvCV60.h2yBfgNjGpc__inline-item-view.r-iCW_xhUARSoI async-local-kp.iCW_xhUARSoI-AjwZTeXvKls.r-iMpgniV_zCxk div.iMpgniV_zCxk-AjwZTeXvKls.y.yf div.immersive-container:nth-child(2) div.akp_tsuid3.akp-el div.t3HED div.g.kno-kp.mnr-c.g-blk div.kp-blk.knowledge-panel.Wnoohf.OJXvsb div.xpdopen div.ifM9O:nth-child(1) div.mod:nth-child(5) div.Z1hOCe div.zloOqf.PZPZlf.kno-fb-ctx span.LrzXr.kno-fv > a.fl:nth-child(1)")).getText();
				row.createCell(7).setCellValue(Affiliation);
				fos = new FileOutputStream(src); 
				wb.write(fos);
			}
			catch(Exception e9)
			{
				
			}
		}
	}
	
	list2.clear();	
	
	k();
	}
public void k() throws InterruptedException, IOException
{
	try
	{
	driver.findElement(By.cssSelector("#pnnext>span")).click();
	   Thread.sleep(3000);
	   m();
	   //h();
	}
	catch(Exception e)
	{
		System.out.println("completed");
	}
}
}
