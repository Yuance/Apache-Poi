package log4j;

import org.apache.log4j.*;

public class MyLogTest {
	static Logger logger = Logger.getLogger(MyLogTest.class.getName());
	
	public static void main1(String[] args) {
		PropertyConfigurator.configure("Log4j.properties");
		MyLogTest.logger.info("Test start.");
		MyLogTest.logger.info("Test ends.");
	}
}