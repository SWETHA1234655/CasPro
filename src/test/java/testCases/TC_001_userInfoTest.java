package testCases;

import org.testng.Assert;
import org.testng.annotations.BeforeClass;

import org.testng.annotations.Test;

import PageObjects.homePage;
import Utilities.ExcelOutputFile;
import testBase.BaseClass;

public class TC_001_userInfoTest extends BaseClass {
	
	// Test method to navigate to the Be Cognizant page and verify the URL.
	@Test(priority=0)
	public void navigateBeCognizantPage() {
		
		hp=new homePage(driver);
		
		String exp_res="https://be.cognizant.com";
		String act=driver.getCurrentUrl();
		if(exp_res.equals(act)) {
			System.out.println("Be.cognizant page is open");
		}
	}
	
	//Test method to verify user information by interacting with the UI elements.
	@Test(priority=1)
	public void verifyUserInfo() throws Exception{
		
		// Log the start of the test case
		logger.info("**** Starting TC_001_UserInfoTest  ****");
		homePage hp=new homePage(driver);
		Thread.sleep(5000);
		
		// Click on the user button to reveal user information
		hp.clickUserButton();
		Thread.sleep(5000);
		
		// Retrieve the user's name from the page
		String name=hp.getUserName();
		System.out.println(name);
		Thread.sleep(5000);
		
		// Retrieve the user's email from the page
		String mail=hp.getUserMail();	
		System.out.println(mail);
		
		// Write the user's name and email to the Excel file
		ExcelOutputFile.setCellData("profile", 1, 0, name);
		ExcelOutputFile.setCellData("profile", 1, 1, mail);
		
		boolean user_result1=name.contains("(Contractor)");
		Assert.assertEquals(user_result1, true);
		
		boolean user_result2=mail.contains("@cognizant.com");
		Assert.assertEquals(user_result2, true);
		
		// Log the end of the test case
		logger.info("**** Finish TC_001_UserInfoTest  ****");
		
		
	}
	

}
