package Generic_utility;

import java.io.File;
import java.time.Duration;
import java.util.Iterator;
import java.util.Set;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;

public class  WebDriver_Utility {
	
	
	/**
	 * This method is used to maximize the window
	 * 
	 * @param driver
	 * @author Shobha
	 */
	public void maximizeWindow(WebDriver driver) {
		driver.manage().window().maximize();
	}

	/**
	 * This method is used, wait elements to get loaded in GUI
	 * 
	 * @param driver
	 */
	public void waitElementsToLoad(WebDriver driver) {
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));

	}

	/**
	 * This method is used for switching windows
	 * 
	 * @param driver
	 * @param partialTitle
	 */
	public void windowSwitching(WebDriver driver, String partialTitle) {
		Set<String> allWins = driver.getWindowHandles();
		Iterator<String> it = allWins.iterator();

		while (it.hasNext()) {
			String win = it.next();
			driver.switchTo().window(win);
			String title = driver.getTitle();
			if (title.contains(partialTitle)) {
				break;
			}

		}
	}

	public void dropDown(WebElement ele, String data) {
		Select select = new Select(ele);
		select.selectByValue(data);
	}

	public void alertHandle(WebDriver driver) {
		driver.switchTo().alert().accept();
	}
	
	public static String takeSCreenShotEx(WebDriver driver,String screenShotName) throws Throwable
	{
		TakesScreenshot ts = (TakesScreenshot)driver;
		File src = ts.getScreenshotAs(OutputType.FILE);
		File dest = new File("./ScreenShots/"+screenShotName+".png");
		FileUtils.copyFile(src, dest);
		return dest.getAbsolutePath();
		
}
	
	
}
