package PageObjects;

import java.time.Duration;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import testBase.BaseClass;

public class homePage extends BasePage{
	
	// Constructor to initialize the driver from the BasePage class
	public homePage(WebDriver driver) {
		super(driver);
		
	}
	
	//Locators
	
	@FindBy(xpath="//div[@class='_8ZYZKvxC8bvw1xgQGSkvvA==']")
	WebElement userButton;
	
	@FindBy(xpath="//div[@class='mectrl_accountDetails']/div[1]")
	WebElement userInfo1;
	
	@FindBy(xpath="//div[@class='mectrl_accountDetails']/div[2]")
	WebElement userInfo2;
	
	//Declaration of Explicit wait
	WebDriverWait wait=new WebDriverWait(driver,Duration.ofSeconds(10));
	//Declaration of Actions class
	Actions act=new Actions(driver);
	BaseClass bc=new BaseClass();
	//Declaration of JavascriptExecutor
	JavascriptExecutor js=(JavascriptExecutor)driver;
	
	// Method to click on the user button
	public void clickUserButton() throws Exception{
		
		// Wait until the user button is visible
		wait.until(ExpectedConditions.visibilityOf(userButton));
		//highlight the user button
		bc.highlightElement(js, userButton);
		// Move to the user button element
		act.moveToElement(userButton).perform();
		// Click on the user button
		act.click().perform();
		Thread.sleep(2000);
		act.doubleClick().perform();
		Thread.sleep(2000);
		
	}
	
	// Method to retrieve the user's name
	public String getUserName() {
		return userInfo1.getText(); 
		
	}
	// Method to retrieve the user's mail
	public String getUserMail() {
		return userInfo2.getText();
	}
	
}
