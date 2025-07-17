package Campaign;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.time.Duration;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.Assert;
import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import Generic_utility.BaseClass;
import Generic_utility.Excel_Utility;
import POM_Repo.CampLookUpImgPage;
import POM_Repo.CreateCampaignPAge;
import POM_Repo.HomePage;
import POM_Repo.ValidationCampaignPage;

//@Listeners(Generic_Utilities.ListenerImp.class)
@Listeners(Generic_utility.ExtentReportImp.class)
public class CreateCampaignsTest extends BaseClass {
	//@Test(retryAnalyzer =Generic_Utilities.RetryAnalyserImp.class )
	@Test
	public void createCampaignsTest() throws Throwable {

		// Login to vtiger application->mouseOverOn more Link->click on campaigns->
		// click on create campaign lookup image->Enter campaignName->
		// click on save Btn->verify whether the campaign is created in campaign
		// Information page
		// and Logout from the application.

		HomePage home = new HomePage(driver);
		home.clickCampaignsLink();

		//pushing 
		//pushing2
		
		CampLookUpImgPage campLookUp = new CampLookUpImgPage(driver);
		campLookUp.clickLookUpImg();

		Excel_Utility elib = new Excel_Utility();
		String campName = elib.getExcelData("Campaign", 0, 0);

		CreateCampaignPAge campPage = new CreateCampaignPAge(driver);
		campPage.enterCampNAme(campName);
		campPage.clickSaveButton();

//		Assert.fail();
		
		ValidationCampaignPage validate = new ValidationCampaignPage(driver);
		validate.validateCamp(driver, campName);

	}
	
//-----------------------------------------------------------------------------------------------------------------------------------
	
}

