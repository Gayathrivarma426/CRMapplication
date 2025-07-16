package Generic_utility;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Parameters;

import POM_Repo.HomePage;
import POM_Repo.LoginPage;



public class BaseClass {

	public WebDriver driver;
	public static WebDriver sdriver;
		
		@BeforeSuite(groups = {"smokeTest","regressionTest"})
		public void BS()
		{
			System.out.println("DataBase connection");
		}
		
		@BeforeTest(groups = {"smokeTest","regressionTest"})
		public void BT()
		{
			System.out.println("Parallel Execution");
		}
		
//		@Parameters("BROWSER")
		@BeforeClass(groups = {"smokeTest","regressionTest"})
//		public void BC(String BROWSER) throws Throwable
//		{
		public void BC() throws Throwable
		{
			//reading keys from properties file
			File_Uitility flib = new File_Uitility();
			String BROWSER = flib.getKeyAndValue("browser");
		
			//reading keys from cmd prompt
//			String BROWSER = System.getProperty("browser");
			
			if (BROWSER.equalsIgnoreCase("chrome")) {
				driver = new ChromeDriver();
			} else if (BROWSER.equalsIgnoreCase("firefox")) {
				driver = new FirefoxDriver();
			} else if (BROWSER.equalsIgnoreCase("edge")) {
				driver = new EdgeDriver();
			} else {
				driver = new FirefoxDriver();
			}
			System.out.println("Browser Launching");
			
			sdriver=driver;
		}
		
//		@Parameters({"URL","USERNAME","PASSWORD"})
		@BeforeMethod(groups = {"smokeTest","regressionTest"})
//		public void BM(String URL,String USERNAME,String PASSWORD  ) throws Throwable
//		{
		public void BM() throws Throwable
		{
			File_Uitility flib = new File_Uitility();
			String BROWSER = flib.getKeyAndValue("browser");
			String URL = flib.getKeyAndValue("url");
			String USERNAME = flib.getKeyAndValue("username");
			String PASSWORD = flib.getKeyAndValue("password");
		
//			String URL = System.getProperty("url");
//			String USERNAME = System.getProperty("username");
//			String PASSWORD = System.getProperty("password");
			driver.get(URL);

			LoginPage login = new LoginPage(driver);
			login.loginToApp(USERNAME, PASSWORD);

			WebDriver_Utility wlib = new WebDriver_Utility();
			wlib.maximizeWindow(driver);
			wlib.waitElementsToLoad(driver);
			
			System.out.println("Login to Application");
		}
		
		@AfterMethod(groups = {"smokeTest","regressionTest"})
		public void AM()
		{
			HomePage home = new HomePage(driver);
			home.logoutApp();
			System.out.println("Logout from Application");
		}
		
		@AfterClass(groups = {"smokeTest","regressionTest"})
		public void AC()
		{
			driver.quit();
			System.out.println("close the Browser");
		}
		@AfterTest(groups = {"smokeTest","regressionTest"})
		public void AT()
		{
			System.out.println("parallel execution close");
		}
		
		@AfterSuite(groups = {"smokeTest","regressionTest"})
		public void AS()
		{
			System.out.println("Close DAtaBAse");
		}
	
}
