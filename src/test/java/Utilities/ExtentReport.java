package Utilities;

import java.awt.Desktop;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.testng.ITestContext;
import org.testng.ITestListener;
import org.testng.ITestResult;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;

import testBase.BaseClass;

public class ExtentReport implements ITestListener{
	
	//Its the UI of the report
	public ExtentSparkReporter sparkReporter;
	//Its populate the common information on the report
	public ExtentReports extent;
	//creating the testcase entries in the report and update status of the test method.
	public ExtentTest test;
	
	BaseClass bc=new BaseClass();
	String reportName;
	
	public void onStart(ITestContext context)
	{
		//To set the timeStamp for the report 
		String timeStamp = new SimpleDateFormat("yyyy.MM.dd.HH.mm.ss").format(new Date());
		reportName = "Test-Report-"+timeStamp+".html";
		//To give the path of the report file
		sparkReporter = new ExtentSparkReporter(System.getProperty("user.dir")+"/reports/"+reportName);
		//To set the title of the document
		sparkReporter.config().setDocumentTitle("Be_Cognizant Featured_News Report ");
		//To set the name of the report
		sparkReporter.config().setReportName("Be_Cognizant Testing");
		//To set the theme of the document
		sparkReporter.config().setTheme(Theme.STANDARD);
		
		//Its to provide some additional information in the report
		extent = new ExtentReports();
		extent.attachReporter(sparkReporter);
		extent.setSystemInfo("Application", "Be_Cognizant");
		extent.setSystemInfo("Module", "Featured News");
		extent.setSystemInfo("User Name", System.getProperty("user.name"));
		extent.setSystemInfo("Environment", "QA");
		
		//Its to select which browser to display the report
		String browser =context.getCurrentXmlTest().getParameter("browser");
		extent.setSystemInfo("Browser", browser);
		
	}
	
	//This method is display the test case is run successfully
	public void onTestSuccess(ITestResult result)
	{
		test=  extent.createTest(result.getName());
		test.assignCategory(result.getMethod().getGroups());
		test.log(Status.PASS, result.getName()+" got successfully executed");
		try {
			String imgPath = bc.screenShot(result.getName());
			test.addScreenCaptureFromPath(imgPath);
		} catch (IOException e1) {
			e1.printStackTrace();
		}
		
	}
	//This method is display the test case is failed
	public void onTestFailure(ITestResult result)
	{
		test = extent.createTest(result.getName());
		test.assignCategory(result.getMethod().getGroups());
		test.log(Status.FAIL,result.getName()+" got failed");
		test.log(Status.INFO,result.getThrowable().getMessage());
		try {
			String imgPath = bc.screenShot(result.getName());
			test.addScreenCaptureFromPath(imgPath);
		} catch (IOException e1) {
			e1.printStackTrace();
		}
	}
	//This method is display the test case is skipped
	public void onTestSkipped(ITestResult result)
	{
		test = extent.createTest(result.getName());
		test.assignCategory(result.getMethod().getGroups());
		test.log(Status.SKIP, result.getName()+" got skipped");
		test.log(Status.INFO, result.getThrowable().getMessage());
	}
	//This method is finish the test case
	public void onFinish(ITestContext context)
	{
		extent.flush();
		String extentReportPath = System.getProperty("user.dir")+"/reports/"+reportName;
		File extentReport = new File(extentReportPath);
		
		try {
			Desktop.getDesktop().browse(extentReport.toURI());
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	
}

	
	


