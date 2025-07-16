package POM_Repo;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import net.bytebuddy.implementation.auxiliary.MethodCallProxy;

public class CreateCampaignPAge {

	public CreateCampaignPAge(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}
	
	@FindBy(name = "campaignname")
	private WebElement campNAme;
	
	@FindBy(xpath = "//input[@title='Save [Alt+S]']")
    private WebElement saveButton;

	public WebElement getCampNAme() {
		return campNAme;
	}

	public WebElement getSaveButton() {
		return saveButton;
	}
	
	/**
	 * This method is used to enter campagin Name
	 * @param name
	 */
	public void enterCampNAme(String name)
	{
		campNAme.sendKeys(name);
	}
	
	/**
	 * This method is used to click on save Button
	 */
	public void clickSaveButton()
	{
		saveButton.click();
	}
}

