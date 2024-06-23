package testBase;

import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.time.Duration;
import java.util.Properties;

import org.apache.commons.io.FileUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.testng.annotations.*;

import PageObjects.homePage;
import Utilities.ExcelOutputFile;
import PageObjects.NewsPage;

public class BaseClass {
	
	public static WebDriver driver;
	public static Logger logger;
	public Properties p;
	public static homePage hp;
	public static NewsPage np;
	public String xlFile=System.getProperty("user.dir") + "\\testData\\outputfile.xlsx";
	
		
	@BeforeClass
	@Parameters({"browser"})
	public void setUp(String browser) throws Exception{
		
		//To read a properties file
		FileReader file = new FileReader("./src/test/resources/config.properties");
		p = new Properties();
		//To load the properties file
		p.load(file);
		
		//Loading log4j2 file
		logger = LogManager.getLogger(this.getClass());
		//choose the browser to run the test cases
		switch(browser.toLowerCase())
		{
		case "chrome": driver = new ChromeDriver(); break;
		case "edge": driver = new EdgeDriver(); break;
		default: System.out.println("No Browsers matching");
		return;
		}
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		
		//To get a url in properties file
		driver.get(p.getProperty("appURL"));
		//maximize the window
		driver.manage().window().maximize();
		// Create an Excel output file
		ExcelOutputFile.createExcel();
	}
	@AfterClass
	public void tearDown()
	{
		driver.quit();
	}
	
	// Captures a screenshot and saves it to a specified location.
	public String screenShot(String testName) throws IOException
	{
		TakesScreenshot t = (TakesScreenshot) driver;
		File sourceFile = t.getScreenshotAs(OutputType.FILE);
		String targetFilePath=System.getProperty("user.dir")+"\\screenshots\\" + testName + ".png";
		File targetFile=new File(targetFilePath);
		sourceFile.renameTo(targetFile);
		return targetFilePath;
		
	}
	//Method to highlight the elements
	public static void highlightElement(JavascriptExecutor jExecutor, WebElement element) {
		jExecutor.executeScript("arguments[0].style.border='3px solid pink'", element);
	}
	
}
