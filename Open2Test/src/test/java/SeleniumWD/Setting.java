package SeleniumWD;
import java.util.List;

import org.openqa.selenium.WebElement;
public class Setting
{
	
	//public String utilityFilePath="<<Give the SeleniumUtility Path here Example:D://SeleniumUtility.xls>>";
	//public String utilityFilePath="E:\\Open 2Test  c folder Back UP\\New folder (3)\\Xlsx_Testing\\O2T_RegTest_Ph11\\O2T_RegTest_Ph11\\TestUtility\\SeleniumUtility_TestO2TOriginalOperation_IE.xlsx";
	public String utilityFilePath="E:\\Open2Test\\Open2Test_Selenium\\SampleScript\\Utility\\SeleniumUtility.xlsx";
	String[] envprevMonth1={"prev","Prev"}; // Specify a class name through which the framework can identify the image representing the previous month
	String[] envnextMonth1={"next","Next"};// Specify a class name through which the framework can identify the image representing the next month
	String[] envtitleMonth={"month"};// Specify a class name through which we can identify the title month element in calendar control element
	String[] envtitleYear={"year"};// Specify a class name through which we can identify the title year element in calendar control element
}
