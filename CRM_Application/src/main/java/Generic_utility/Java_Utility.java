package Generic_utility;

import java.util.Random;

public class Java_Utility {

	
	/**
	 * This Method is used to avoid duplicate values
	 * 
	 * @return
	 * @author Shobha
	 */
	public int getRandomNum() {
		Random ran = new Random();
		int ranNum = ran.nextInt(1000);
		return ranNum;
	}
}
