package testCases;

import java.io.IOException;

import java.util.List;

import org.openqa.selenium.WebElement;
import org.testng.Assert;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import PageObjects.NewsPage;
import testBase.BaseClass;

public class TC_002_NewsPageTest extends TC_001_userInfoTest {
	
	// Test method to verify that feature news elements are displayed on the NewsPage.	
	@Test(priority=2)
	public void verifynewsDisplayed() throws Exception {
		
		logger.info("*********   TC_001_Feature news page is display     ************* ");
		np=new NewsPage(driver);
		int count=0;
			List<WebElement>news=np.getAllNewsElements();
			for(WebElement n:news) 
			{
				n.isDisplayed();
				count++;
			}
			if(count==5) {
				Assert.assertTrue(true);
			}
			else {
				Assert.assertTrue(false);
			}
		logger.info("********* TC_001_Feature news displyed successfully   ********");	
	}

	//Test method to retrieve and verify details of each feature news item.
	@Test(priority=3)
	public void getNewsPageDetails() throws Exception
	{
		np=new NewsPage(driver);
		
		List<WebElement> featureNews = np.getAllNewsElements();
		
		for(int i=0;i<featureNews.size();i++)
		{
			List<WebElement> featureNews1 = np.getAllNewsElements();
			
			featureNews.get(i).click();
			
			logger.info("********* TC_002_verifyToolTip ***********");
			
			np.verifyTitleTooltip(i+1);
			
			logger.info("********* TC_003_verifyHoverAuthorDetails *********");
			
			np.hoverAuthorDetails();
			
			logger.info("********* TC_004_shareOptions *********");
			
			np.clickShareButton();
			
			np.getShareOptions();
			
			logger.info("********* TC_005_verifyNewsLetterLink Display*********");
			
			np.getNewsletterLink();
			
			logger.info("********* TC_006_verifyLikes Display *********");
			
			np.totalLikes();
			
			logger.info("********* TC_007_verifyViews Display *********");
			
			np.totalViews();
			
			Thread.sleep(5000);
			
			np.navigate();
			
			System.out.println("===================================================");
			
			logger.info("******** News details are displayed successfully ******");
			
		}
	}
}
