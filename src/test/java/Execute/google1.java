package Execute;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.atomic.AtomicInteger;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.Proxy;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class google1 
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
	String Founded = "";
	String Affiliation = "";
	AtomicInteger count = new AtomicInteger(1);
	@BeforeTest
	public void launch() throws InterruptedException, IOException
	{
		Proxy p=new Proxy();
	    p.setHttpProxy("45.77.156.187:8080");
	    DesiredCapabilities cap=new DesiredCapabilities();     
	      // Pass proxy object p
	    cap.setCapability(CapabilityType.PROXY, p);
	   System.setProperty("webdriver.chrome.driver", "D:\\Software\\Selenium\\New Driver\\Chrome Driver 83\\chromedriver.exe");
	   System.setProperty("webdriver.chrome.silentOutput", "true");
	    driver = new ChromeDriver();
		//FirefoxOptions foptions = new FirefoxOptions();
	  // foptions.setCapability("marionette", false);
	   // driver = new FirefoxDriver(foptions);
	    driver.manage().window().maximize(); 
	    src = new File("D:\\Software\\Database_2020\\db shop database\\allindiaschooldatabase (1)\\May 2020\\Google\\schoolurl3.xlsx");
		wb = new XSSFWorkbook(); 
		sheet = wb.createSheet("wp_details");	
		row = sheet.createRow((short)0);  
		row.createCell(0).setCellValue("city");
		row.createCell(1).setCellValue("url");
	    fos = new FileOutputStream(src); 
		wb.write(fos);
		driver.navigate().to("https://www.google.com/search?tbm=lcl&ei=9xIIXK3rBNey9QPIl6fIBA&q=furniture+dealer&oq=furniture+dealer&gs_l=psy-ab.3..35i39k1j0l9.1603425.1614009.0.1614899.16.16.0.0.0.0.247.2142.0j15j1.16.0....0...1c.1.64.psy-ab..0.16.2137...0i131k1j0i67k1.0.GAisbOSxAYE#rlfi=hd:;si:;mv:!1m2!1d28.667845699999997!2d77.4001945!2m2!1d28.5082693!2d77.1836514;tbs:lrf:!2m1!1e2!2m1!1e3!3sIAE,lf:1,lf_ui:10");
		Thread.sleep(30000);
	}
	@Test
	public void a() throws IOException, InterruptedException
	{
		for(int i =25; i<30; i++)
		{
			String city = Readwriteexcel.readexcel("D:\\Software\\Database_2020\\db shop database\\allindiaschooldatabase (1)\\May 2020\\Google\\city_input_main.xlsx", 0, i, 0);
			for(int j =1; j<8; j++)
			{
		    String category = Readwriteexcel.readexcel("D:\\Software\\Database_2020\\db shop database\\allindiaschooldatabase (1)\\May 2020\\Google\\category_input_new.xlsx", 0, j, 1);
			String input = category+ " " +"in"+ " "+city;
			driver.findElement(By.xpath("//*[@id=\"lst-ib\"]")).clear();
			driver.findElement(By.xpath("//*[@id=\"lst-ib\"]")).sendKeys(input);
			Actions act = new Actions(driver); 
			act.sendKeys(Keys.ENTER).build().perform(); 
			Thread.sleep(3000);
			String requiredurl = driver.getCurrentUrl();
			row = sheet.createRow((short)count.getAndIncrement());  
			row.createCell(0).setCellValue(city);
			row.createCell(1).setCellValue(requiredurl);
			row.createCell(2).setCellValue(category);
		    fos = new FileOutputStream(src); 
			wb.write(fos);
			}
		}
	}
	
}