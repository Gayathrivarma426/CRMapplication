package POM_Repo;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.PageFactory;

public class ValidateOrgPage {

	
	public ValidateOrgPage(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}
	
	public void validateOrg(WebDriver driver,String name)
	{
		String actdata = driver.findElement(By.id("dtlview_Organization Name")).getText();
		if (actdata.contains(name)) {
			System.out.println("Organization Name is Created");
		} else {

			System.out.println("Organization Name is not Created");
		}
	}
}
