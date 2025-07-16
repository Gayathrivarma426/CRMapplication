package Generic_utility;

import java.io.FileInputStream;
import java.util.Properties;

public class File_Uitility {

	
	/**
	 * This method is used to read data from properties file(External resource)
	 * 
	 * @param key
	 * @return
	 * @throws Throwable
	 * @author Shobha
	 */
	public String getKeyAndValue(String key) throws Throwable {
		FileInputStream fis = new FileInputStream("C:\\Users\\sures\\OneDrive\\Desktop\\AdvanceSelenium7AM\\FileUtilityCommanData.properties");

		// step2:- Create the object of Properties class and load all the Keys
		Properties pro = new Properties();
		pro.load(fis);

		// step3:-read the value using getProperty()
		String Value = pro.getProperty(key);

		return Value;

	}
	
}
