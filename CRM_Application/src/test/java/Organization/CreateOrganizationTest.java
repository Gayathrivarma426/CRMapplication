package Organization;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.time.Duration;
import java.util.Properties;
import java.util.Random;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.Test;

import Generic_utility.BaseClass;
import Generic_utility.Excel_Utility;
import Generic_utility.Java_Utility;
import POM_Repo.HomePage;
import POM_Repo.OrganizationCreatePage;
import POM_Repo.OrganizationLookUpImg;
import POM_Repo.ValidateOrgPage;



public class CreateOrganizationTest extends BaseClass {

	@Test(groups = {"smokeTest","regressionTest"})
	public void createOrganizationTest() throws Throwable {

		// Login to vtiger application
		// ->click on organizations link
		// ->click on create organization lookup image
		// ->Enter organisation name,phone number and email
		// ->click on save Btn
		// ->verify whether the organization is created in Organization Information page
		// and
		// Logout from the application.

		

		HomePage home = new HomePage(driver);
		home.clickOrgLink();

		OrganizationLookUpImg lookUp = new OrganizationLookUpImg(driver);
		lookUp.clickPrdLookUp();

		Java_Utility jlib = new Java_Utility();
		int ranNum = jlib.getRandomNum();

		Excel_Utility elib = new Excel_Utility();
		String OrgName = elib.getExcelData("Organization", 0, 0) + ranNum;
		String PhnNum = elib.getExcelDataUsingFormatter("Organization", 1, 0);
		String mailId = elib.getExcelDataUsingFormatter("Organization", 2, 0);

		OrganizationCreatePage orgPage = new OrganizationCreatePage(driver);
		orgPage.enterOrgnaizationData(OrgName, PhnNum, mailId);

		ValidateOrgPage validate = new ValidateOrgPage(driver);
		validate.validateOrg(driver, OrgName);

	}
	//----------------------------------------------------------------------------------------------------


}
