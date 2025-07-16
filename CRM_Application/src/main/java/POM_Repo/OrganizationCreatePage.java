package POM_Repo;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import net.bytebuddy.asm.MemberSubstitution.FieldValue;

public class OrganizationCreatePage {

	public OrganizationCreatePage(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}
	
	@FindBy(name="accountname")
	private WebElement orgName;
	
	@FindBy(name = "phone")
	private WebElement phnNum;
	
	@FindBy(id = "email1")
	private WebElement mailId;

	@FindBy(css = "[title='Save [Alt+S]']")
	private WebElement saveButton;
	
	public WebElement getOrgName() {
		return orgName;
	}

	public WebElement getPhnNum() {
		return phnNum;
	}

	public WebElement getMailId() {
		return mailId;
	}
	/**
	 * 
	 * @param name
	 * @param PhoneNum
	 * @param email
	 */
	public void enterOrgnaizationData(String name,String PhoneNum,String email)
	
	{
		orgName.sendKeys(name);
		phnNum.sendKeys(PhoneNum);
		mailId.sendKeys(email);
		saveButton.click();
	}
	
	
}

