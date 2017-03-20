package SeleniumWD;

/*######################################################################################################
 'Project Name		: Open2Test framework for Selenium Web Driver
 'Author		    : Open2Test
 'Version	    	: V 3.1.0
 'Date of Creation	: 28-March-2016
 '######################################################################################################
 */
//V2- Enhanced with context support
//V3 - Enahanced with FileUploadDownload Management
import static org.junit.Assert.*;

import org.junit.After;
import org.junit.Test;


//import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert; //Updated on 16.12.2013 for Dialog support
import org.openqa.selenium.By;
import org.openqa.selenium.Capabilities;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoAlertPresentException;//Updated on 16.12.2013 for Dialog support
import org.openqa.selenium.OutputType;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.remote.Augmenter;
import org.openqa.selenium.remote.RemoteWebDriver;

import java.io.*;

import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.Select;

import ru.yandex.qatools.ashot.AShot;
import ru.yandex.qatools.ashot.Screenshot;


 


//import ru.yandex.qatools.ashot.screentaker.ViewportPastingStrategy;
import ru.yandex.qatools.ashot.shooting.ShootingStrategies;
import ru.yandex.qatools.ashot.shooting.ShootingStrategy;





import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet; //Updated on 16.12.2013 for page support
import java.util.Iterator; //Updated on 16.12.2013 for page support
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set; //Updated on 16.12.2013 for page support
import java.util.Stack;
import java.util.concurrent.TimeUnit;
import java.util.logging.ConsoleHandler;
import java.util.logging.FileHandler;
import java.util.logging.Handler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.lang.System;
import java.lang.reflect.Method;
import java.net.InetAddress;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit; // Updated 11.10.2014 screenshots
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;

import javax.imageio.ImageIO;

public class SeleniumWD {
	int scriptfound = 0;
	Setting setObj = new Setting(); // Updated on 16.12.2013
	
	static BufferedWriter bw = null;
	static BufferedWriter bw1 = null;
	static RemoteWebDriver D8;
	String browserver = null;
	String mm = null;
	String dd = null;
	String yyyy = null;
	String monthpart;
	String monthpartjap;
	Date cur_dt = null;
	String testSuite;
	String testScript;
	
	String database;
	String host_name;
	String portnumber;
	String schemaname;
	String username;
	String password;
	String sqlquery;
	Connection con;
	Statement stmt;
	ResultSet rs;
	String update=null;

	static XSSFSheet ws;
	boolean isinvaliddb = false;
	boolean isconnected =false;
	
	static String globalCompName;
	static String objectRepository;
	int startrow = 0;
	static String reportsPath;
	String strResultPath;
	String[] TCNm = null;
	
	static String exeStatus_suite="True";
	
	static String exeStatus = "True";
	static String screenReport;
	static int rowcnt;
	static int ORrowcount = 0;
	static String ORvalname = "";
	static int dtrownum = 1;
	static String ORvalue = "";
	static String Action = "";
	static String cCellData = "";
	static String dCellData = "";
	static String sComments="";
	static double nComments=0;
	static boolean  bComments=false;
	static String htmlname = "";
	static String cCellObjType = "";// included on 29-Nov-2013 for storing
	static String cCellObjName = ""; // included on 29-Nov-2013 for storing
										// object name
	String[] cCellDataVal = null;
	String[] dCellDataVal = null;
	String ObjectSet;
	String[] ObjectSet1 = null;
	String ObjectSet2 = "";
	String ObjectSetVal = "";
	static XSSFSheet DTsheet = null;
	static XSSFSheet ORsheet;
	String Searchtext;
	String parentWindowHandle;// Updated on 16.12.2013 for page support
	String parentWindowHandle1;
	String windowHandle;// Updated on 16.12.2013 for page support
	String DriverPath;// Updated on 16.12.2013 for outside browser setting
	String tempTestReportPath;// Updated on 16.12.2013 for outside temporary
								// report path
	String ieDriverPath;
	String chromeDriverPath;
	String ffDriverPath;
	String reportsFolder="";
	static int iflag = 0;
	static String comments =null;

	static int screenshotflag = 0;
	static int loopflag = 0;
	static int j = 0;
	static int loopsize = -1;
	int[] loopstart = new int[1];
	int[] loopcount = new int[1];
	int[] loopend = new int[1];
	static int[] loopcnt = new int[1];
	static int[] dtrownumloop = new int[1];
	boolean captureperform = false;
	static boolean capturecheckvalue = false; 
	static boolean capturestorevalue = false;
	static XSSFSheet TScsheet;
	static XSSFWorkbook TScworkbook;
	static String execpath;
	static int TScrowcount = 0;
	static int loopnum = 1;
	static int windowFound = 0;
	static String TScname;
	static String ActionVal;
	String BrowserType;// Assign with either FF or IE
	static int DTcolumncount = 0;
	static WebElement elem = null;
	Alert dialogSwitch = null; // Updated on 20-Dec-2013
	// Updated on on 16.12.2013
	private static Map<String, String> map = new HashMap<String, String>();
	// private static Map<String, Float> mapint = new HashMap<String, Float>();
	public static int devFlag = 0;
	public static String varname;
	public static int initscreen = 1;
	public static int initscreenFlag = 1;
	String browserName;
	public static String objectType;
	public static int objFoundFlag = 0;
	public static String capString = "";
	static int locatorError = 0;
	static int ScreenshotTypeFlag = 0;
	int varUpdateReport;
	int today = 0;
	String selectedDateClass = "";
	int lastSelecteddateJQ = 0;
	int lastSelectedmonthJQ = 0;
	int lastSelectedyearJQ = 0;
	static int conditionline = 0;
	static int reporttype = 0;
	static String reusableComponents;
	static String ExpectedReportData = "";
	static String ActualReportData = "";
	static String OptionalString = "";
	static String OptionalString2 = "";
	static String[] getCellArray;
	static int cCellDatanum =0;
	static int waitcCellDatanum=0;
	static String DetailReport="";
	static String ScreenshotPath="";
	public static String ScreenFolder;
	static int browsertypeflg =0;
	static String browser;
	String ScriptName;
	
	private static final Logger LOGGER = Logger.getLogger(SeleniumWD.class.getName());

	/*
	 * This function reads the selenium utility file and identifies where Object
	 * Repository, Test Suite & Test Scripts are located
	 */
	//@SuppressWarnings("unchecked")
	@SuppressWarnings({"unchecked", "deprecation", "resource"})
	@Test
	public void readUtilFile() {
		   LoadLog();
		XSSFWorkbook w1 = null;
		try {
			 
			if (!setObj.utilityFilePath.toString().endsWith("xlsx")) {
				System.out
						.println("Utility file defined in Setting.Java is not a correct one. Give the correct (xlsx) Format of Utility File");
				LOGGER.log(Level.SEVERE, "Error occur in readUtilFile.","Utility file defined in Setting.Java is not a correct one. Give the correct (xlsx) Format of Utility File");
				return;
			} else {
				w1 = new XSSFWorkbook(setObj.utilityFilePath);
				 
			}

		} catch (Exception e) {
			System.out
					.println("Selenium Utility file is not accessible. Check your Utility file in the given path  "+e.getMessage());
			LOGGER.log(Level.SEVERE, "Error occur in readUtilFile.","Selenium Utility file is not accessible. Check your Utility file in the given path  ");
			return;
		} 
		
		
		try {
			XSSFSheet sheet = w1.getSheetAt(0);
			
			if(sheet.getRow(1).getCell(1).getCellType()==XSSFCell.CELL_TYPE_BLANK)	{
				testSuite ="";
			}else if(sheet.getRow(1).getCell(1).getCellType()==XSSFCell.CELL_TYPE_NUMERIC){
			double	password1 = sheet.getRow(1).getCell(1).getNumericCellValue();
			testSuite = String.valueOf(password1);
			}else if(sheet.getRow(1).getCell(1).getCellType()==XSSFCell.CELL_TYPE_STRING){
				testSuite = sheet.getRow(1).getCell(1).getStringCellValue();
			}else if(sheet.getRow(1).getCell(1).getCellType()==XSSFCell.CELL_TYPE_FORMULA){
				testSuite = sheet.getRow(1).getCell(1).getStringCellValue();
			}
			
		
			if(sheet.getRow(2).getCell(1).getCellType()==XSSFCell.CELL_TYPE_BLANK)	{
				testScript ="";
			}else if(sheet.getRow(2).getCell(1).getCellType()==XSSFCell.CELL_TYPE_NUMERIC){
			double	password1 = sheet.getRow(2).getCell(1).getNumericCellValue();
			testScript = String.valueOf(password1);
			}else if(sheet.getRow(2).getCell(1).getCellType()==XSSFCell.CELL_TYPE_STRING){
				testScript = sheet.getRow(2).getCell(1).getStringCellValue();
			}else if(sheet.getRow(2).getCell(1).getCellType()==XSSFCell.CELL_TYPE_FORMULA){
				testScript = sheet.getRow(2).getCell(1).getStringCellValue();
			}	 
			 
			
			if(sheet.getRow(3).getCell(1).getCellType()==XSSFCell.CELL_TYPE_BLANK)	{
				objectRepository ="";
			}else if(sheet.getRow(3).getCell(1).getCellType()==XSSFCell.CELL_TYPE_NUMERIC){
			double	password1 = sheet.getRow(3).getCell(1).getNumericCellValue();
			objectRepository = String.valueOf(password1);
			}else if(sheet.getRow(3).getCell(1).getCellType()==XSSFCell.CELL_TYPE_STRING){
				objectRepository = sheet.getRow(3).getCell(1).getStringCellValue();
			}else if(sheet.getRow(3).getCell(1).getCellType()==XSSFCell.CELL_TYPE_FORMULA){
				objectRepository = sheet.getRow(3).getCell(1).getStringCellValue();
			}			
		 
			
		 
			if(sheet.getRow(5).getCell(1).getCellType()==XSSFCell.CELL_TYPE_BLANK)	{
				reportsPath ="";
			}else if(sheet.getRow(5).getCell(1).getCellType()==XSSFCell.CELL_TYPE_NUMERIC){
			double	password1 = sheet.getRow(5).getCell(1).getNumericCellValue();
			reportsPath = String.valueOf(password1);
			}else if(sheet.getRow(5).getCell(1).getCellType()==XSSFCell.CELL_TYPE_STRING){
				reportsPath = sheet.getRow(5).getCell(1).getStringCellValue();
			}else if(sheet.getRow(5).getCell(1).getCellType()==XSSFCell.CELL_TYPE_FORMULA){
				reportsPath = sheet.getRow(5).getCell(1).getStringCellValue();
			}				 
			 
			if(sheet.getRow(6).getCell(1).getCellType()==XSSFCell.CELL_TYPE_BLANK)	{
				screenReport ="";
			}else if(sheet.getRow(6).getCell(1).getCellType()==XSSFCell.CELL_TYPE_NUMERIC){
			double	password1 = sheet.getRow(6).getCell(1).getNumericCellValue();
			screenReport = String.valueOf(password1);
			}else if(sheet.getRow(6).getCell(1).getCellType()==XSSFCell.CELL_TYPE_STRING){
				screenReport = sheet.getRow(6).getCell(1).getStringCellValue();
			}else if(sheet.getRow(6).getCell(1).getCellType()==XSSFCell.CELL_TYPE_FORMULA){
				screenReport = sheet.getRow(6).getCell(1).getStringCellValue();
			}	
			
			
			if(sheet.getRow(7).getCell(1).getCellType()==XSSFCell.CELL_TYPE_BLANK)	{
				browserName ="";
			}else if(sheet.getRow(7).getCell(1).getCellType()==XSSFCell.CELL_TYPE_NUMERIC){
			double	password1 = sheet.getRow(7).getCell(1).getNumericCellValue();
			browserName = String.valueOf(password1);
			}else if(sheet.getRow(7).getCell(1).getCellType()==XSSFCell.CELL_TYPE_STRING){
				browserName = sheet.getRow(7).getCell(1).getStringCellValue();
				
			 
				
			}else if(sheet.getRow(7).getCell(1).getCellType()==XSSFCell.CELL_TYPE_FORMULA){
				browserName = sheet.getRow(7).getCell(1).getStringCellValue();
			}
			 
			
			if(sheet.getRow(9).getCell(1).getCellType()==XSSFCell.CELL_TYPE_BLANK)	{
				tempTestReportPath ="";
			}else if(sheet.getRow(9).getCell(1).getCellType()==XSSFCell.CELL_TYPE_NUMERIC){
			double	password1 = sheet.getRow(9).getCell(1).getNumericCellValue();
			tempTestReportPath = String.valueOf(password1);
			}else if(sheet.getRow(9).getCell(1).getCellType()==XSSFCell.CELL_TYPE_STRING){
				tempTestReportPath = sheet.getRow(9).getCell(1).getStringCellValue();
			}else if(sheet.getRow(9).getCell(1).getCellType()==XSSFCell.CELL_TYPE_FORMULA){
				tempTestReportPath = sheet.getRow(9).getCell(1).getStringCellValue();
			}
			
			 
			if(sheet.getRow(8).getCell(1).getCellType()==XSSFCell.CELL_TYPE_BLANK)	{
				DriverPath ="";
			}else if(sheet.getRow(8).getCell(1).getCellType()==XSSFCell.CELL_TYPE_NUMERIC){
			double	password1 = sheet.getRow(8).getCell(1).getNumericCellValue();
			DriverPath = String.valueOf(password1);
			}else if(sheet.getRow(8).getCell(1).getCellType()==XSSFCell.CELL_TYPE_STRING){
				DriverPath = sheet.getRow(8).getCell(1).getStringCellValue();
				if (!DriverPath.endsWith("\\")) { // checks whether the path ends with \
				 
					DriverPath = DriverPath + "\\";
                                 }
			
			}else if(sheet.getRow(8).getCell(1).getCellType()==XSSFCell.CELL_TYPE_FORMULA){
				DriverPath = sheet.getRow(8).getCell(1).getStringCellValue();
			}
			
			if(sheet.getRow(11).getCell(1).getCellType()==XSSFCell.CELL_TYPE_BLANK)	{
				database ="";
			}else if(sheet.getRow(11).getCell(1).getCellType()==XSSFCell.CELL_TYPE_NUMERIC){
			double	password1 = sheet.getRow(11).getCell(1).getNumericCellValue();
			database = String.valueOf(password1);
			}else if(sheet.getRow(11).getCell(1).getCellType()==XSSFCell.CELL_TYPE_STRING){
				database = sheet.getRow(11).getCell(1).getStringCellValue();
			}else if(sheet.getRow(11).getCell(1).getCellType()==XSSFCell.CELL_TYPE_FORMULA){
				database = sheet.getRow(11).getCell(1).getStringCellValue();
			}
			
			if(sheet.getRow(12).getCell(1).getCellType()==XSSFCell.CELL_TYPE_BLANK)	{
				host_name ="";
			}else if(sheet.getRow(12).getCell(1).getCellType()==XSSFCell.CELL_TYPE_NUMERIC){
			double	password1 = sheet.getRow(12).getCell(1).getNumericCellValue();
			host_name = String.valueOf(password1);
			}else if(sheet.getRow(12).getCell(1).getCellType()==XSSFCell.CELL_TYPE_STRING){
				host_name = sheet.getRow(12).getCell(1).getStringCellValue();
			}else if(sheet.getRow(12).getCell(1).getCellType()==XSSFCell.CELL_TYPE_FORMULA){
				host_name = sheet.getRow(12).getCell(1).getStringCellValue();
			}
			
		 
			if(sheet.getRow(13).getCell(1).getCellType()==XSSFCell.CELL_TYPE_BLANK)	{
				portnumber ="";
			}else if(sheet.getRow(13).getCell(1).getCellType()==XSSFCell.CELL_TYPE_NUMERIC){
			double	password1 = sheet.getRow(13).getCell(1).getNumericCellValue();
			portnumber = String.valueOf(password1);
			}else if(sheet.getRow(13).getCell(1).getCellType()==XSSFCell.CELL_TYPE_STRING){
				portnumber = sheet.getRow(13).getCell(1).getStringCellValue();
			}else if(sheet.getRow(13).getCell(1).getCellType()==XSSFCell.CELL_TYPE_FORMULA){
				portnumber = sheet.getRow(13).getCell(1).getStringCellValue();
			}
			
			if(sheet.getRow(14).getCell(1).getCellType()==XSSFCell.CELL_TYPE_BLANK)	{
				schemaname ="";
			}else if(sheet.getRow(14).getCell(1).getCellType()==XSSFCell.CELL_TYPE_NUMERIC){
			double	password1 = sheet.getRow(14).getCell(1).getNumericCellValue();
			schemaname = String.valueOf(password1);
			}else if(sheet.getRow(14).getCell(1).getCellType()==XSSFCell.CELL_TYPE_STRING){
				schemaname = sheet.getRow(14).getCell(1).getStringCellValue();
			}else if(sheet.getRow(14).getCell(1).getCellType()==XSSFCell.CELL_TYPE_FORMULA){
				schemaname = sheet.getRow(14).getCell(1).getStringCellValue();
			}
			
		 
			if(sheet.getRow(15).getCell(1).getCellType()==XSSFCell.CELL_TYPE_BLANK)	{
				username ="";
			}else if(sheet.getRow(15).getCell(1).getCellType()==XSSFCell.CELL_TYPE_NUMERIC){
			double	password1 = sheet.getRow(15).getCell(1).getNumericCellValue();
			username = String.valueOf(password1);
			}else if(sheet.getRow(15).getCell(1).getCellType()==XSSFCell.CELL_TYPE_STRING){
				username = sheet.getRow(15).getCell(1).getStringCellValue();
			}else if(sheet.getRow(15).getCell(1).getCellType()==XSSFCell.CELL_TYPE_FORMULA){
				username = sheet.getRow(15).getCell(1).getStringCellValue();
			}
			
		 
		if(sheet.getRow(16).getCell(1).getCellType()==XSSFCell.CELL_TYPE_BLANK)	{
			password ="";
		}else if(sheet.getRow(16).getCell(1).getCellType()==XSSFCell.CELL_TYPE_NUMERIC){
		double	password1 = sheet.getRow(16).getCell(1).getNumericCellValue();
		password = String.valueOf(password1);
		}else if(sheet.getRow(16).getCell(1).getCellType()==XSSFCell.CELL_TYPE_STRING){
			password = sheet.getRow(16).getCell(1).getStringCellValue();
		}else if(sheet.getRow(16).getCell(1).getCellType()==XSSFCell.CELL_TYPE_FORMULA){
			password = sheet.getRow(16).getCell(1).getStringCellValue();
		}
			 
			update = sheet.getRow(17).getCell(1).getStringCellValue();
			 	
			try {
				
				execpath = sheet.getRow(10).getCell(1).getStringCellValue();
				 
				reusableComponents =sheet.getRow(18).getCell(1).getStringCellValue();
				 

			} catch (ArrayIndexOutOfBoundsException exp10) {
				System.out
						.println("WARNING: If you want to use file upload/ download automation feature, use the new format of Utility(Configuration) Excel Which contains the path for the FileManager EXE. ");
				LOGGER.log(Level.SEVERE, "WARNING occur in readUtilFile.","If you want to use file upload/ download automation feature, use the new format of Utility(Configuration) Excel Which contains the path for the FileManager EXE.");
			}
			if (!testSuite.endsWith(".xlsx")) {
				if (testSuite.toString() == "") {
					System.out
							.println("No proper TestSuite is defined in Selenium Utility file. Give the correct TestSuite name.");
				} else {
					System.out
							.println("Invalid File Format of TestSuite in Selenium Utility file: "
									+ testSuite + ". Give only xlsx file format");
					LOGGER.log(Level.SEVERE, "Error occur in readUtilFile.","Invalid File Format of TestSuite in Selenium Utility file: "+ testSuite + ". Give only xlsx file format  ");
				}

			}

			if (!objectRepository.endsWith(".xlsx")) {
				if (objectRepository.toString() == "") {
					System.out
							.println("No proper Object repository is defined in Selenium Utility excel. Give the correct Object repository. ");
					LOGGER.log(Level.SEVERE, "Error occur in readUtilFile.","No proper Object repository is defined in Selenium Utility excel. Give the correct Object repository.");
				} else {
					System.out
							.println("Invalid File Format of Object Repository : "
									+ objectRepository
									+ " in Selenium Utility file. Give only xlsx file format");
					LOGGER.log(Level.SEVERE, "Error occur in readUtilFile.","Invalid File Format of Object Repository : "
							+ objectRepository
							+ " in Selenium Utility file. Give only xlsx file format");
				}

			}
			if (browserName.toString() == "") {
				System.out
						.println("The browser type in the Selenium Utility file is not proper. Give a valid browser type");
				LOGGER.log(Level.SEVERE, "The browser type in the Selenium Utility file is not proper. Give a valid browser type");
			}

			for (int z = 0; z < 1; z++) {
				loopstart[z] = 0;
				loopend[z] = 0;
				loopcnt[z] = 0;
				dtrownumloop[z] = 1;
				loopcount[z] = 0;
			}

		} catch (ArrayIndexOutOfBoundsException ex) {
			System.out
					.println("The format of Utility file is not a correct one. Give the correct (Xlsx) Format of Utility File.");
			LOGGER.log(Level.SEVERE, "The format of Utility file is not a correct one. Give the correct (Xlsx) Format of Utility File");
			System.out.println(ex.getMessage());
			return;
		} catch (Exception e) {
			System.out.println(e.getMessage());
			LOGGER.log(Level.SEVERE, e.getMessage());
			
		}
		try {
			
			String[] browserlist = browserName.split(",");
			for(int i=0;i<browserlist.length;i++){
				
				setBrowser(browserlist[i]);
				findExecTestscript(testSuite, testScript, objectRepository);
				 close();
			}		


		} catch (Exception e) {
			System.out
					.println("Could not launch the Browser. Check the browser settings");
			
			LOGGER.log(Level.SEVERE, "Could not launch the Browser. Check the browser settings");
			return;
		}
		try {
		//	findExecTestscript(testSuite, testScript, objectRepository);

		} catch (Exception e) {
			System.out.println("Error Occured in FindExecTestscript function "+e.getMessage());
			LOGGER.log(Level.SEVERE, "Error Occured in FindExecTestscript function "+e.getMessage());
			
		}
		
	}

	private void LoadLog() {

		Handler consoleHandler = null;
		
        Handler fileHandler  = null;

        try{

       

            consoleHandler = new ConsoleHandler();

            fileHandler  = new FileHandler("./SeleniumWD.log");

             
      

            LOGGER.addHandler(fileHandler);
           
         

            fileHandler.setLevel(Level.ALL);

            LOGGER.setLevel(Level.ALL);        

            LOGGER.config("Configuration done.");
        
             
       LOGGER.log(Level.FINE, "Finer logged");

        }catch(IOException exception){

         LOGGER.log(Level.SEVERE, "Error occur in FileHandler.", exception);

        }
        
			
	}
	
	private  String createConnection(String database, String schemaName, String username, 
			String password) {
	
		String url1=null;
		try {
		if(database.equalsIgnoreCase("mysql")){
			
			String hostadd=  host_name.trim();
			
				InetAddress ipaddr =InetAddress.getByName(host_name.trim());
			 
					
				Class.forName("com.mysql.jdbc.Driver");
				 
				String url = "jdbc:mysql://"+host_name+":"+portnumber+"/"+schemaName;
				 
					if (captureperform == true) {
						screenShot(loopnum, TScrowcount, TScname);
					}
					 
					return url;
		}
					else if(database.equalsIgnoreCase("mssql")){
					
						Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
						
						url1 = "jdbc:sqlserver://";
					
					
						String hostname =host_name;
						String db = "databasename="+schemaName;
						String user = "user="+username;
						String pass  ="password="+password;
		                 String url =url1+host_name+";"+db+";"+user+";"+pass+";";
		                 
		               return url;
						
					}
					else{
						System.out.println(database+" is not handeled as database in this version of  framework");
						update_Report("missing");
						if (captureperform == true) {
							screenShot(loopnum, TScrowcount, TScname);
						}
						return url1;
					}
		
			} catch (Exception e) {
				try {
					update_Report("invalidConnection");
					if (captureperform == true) {
						screenShot(loopnum, TScrowcount, TScname);
					}
				} catch (IOException e1) {
					System.out.println("Invalid Connection  "+e1.getMessage());
					LOGGER.log(Level.SEVERE, "Invalid Connection  "+e1.getMessage());
				 
				} catch (Exception e2) {
				 System.out.println("Error Occured in update_Report Function "+e2.getMessage());
				 LOGGER.log(Level.SEVERE, "Error Occured in update_Report Function "+e2.getMessage());
				}
			
			} 
		
		return url1;
	}
		
	private boolean  getConnection(String url){
		boolean parameter = false;
		SQLException errorString =null;
		try {
			
			if(database.equalsIgnoreCase("mssql")){
			 
			con=DriverManager.getConnection(url);
		 
			
			stmt=con.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE, ResultSet.CONCUR_UPDATABLE);
			
			
			rs=stmt.executeQuery(sqlquery);
			isconnected = true;
			return isconnected;
			
			}else if(database.equalsIgnoreCase("mysql")){
				System.out.println("In get connection part of MSSQL");
				con=DriverManager.getConnection(url, username, password);
				stmt=con.createStatement();
				rs=stmt.executeQuery(sqlquery);
				isconnected = true;
				return isconnected;
			}
			return isconnected;
			
		} catch (SQLException e) {
			
			errorString = e;
			LOGGER.log(Level.SEVERE, "Connection to Database is failed, please verify connection URL parameters which are set in Selenium Utility File  "+e.getMessage());
			System.out.println("Connection to Database is failed, please verify connection URL parameters which are set in Selenium Utility File  "+e.getMessage());
			return isconnected;
	
				
		}
		
	
	}
		
	 private void executeQuery(ResultSet rs){
			 
			String resultfile =null;
			
			 File file = new File("Temp");
			 boolean result=false; 
			 
			   if(!file.exists()){
			    	System.out.println("Creating Directory - "+file);;
			    	try{
			    		file.mkdir();
			    		result = true;
			    	}catch(SecurityException se){
			    	System.out.println(se.getMessage());
			    	LOGGER.log(Level.SEVERE, "SecurityException Error Occured in executeQuery  "+se.getMessage());
			    	}
			    }
			    if(result){
			    	System.out.println("Direcotry is already Created");
			    	LOGGER.log(Level.SEVERE, "Direcotry is already Created");
			    }
			    	
			
			String workingDir = System.getProperty("user.dir");
			 
			 DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
			 Date date = new Date();
			 String CurrentDate= dateFormat.format(date);
		     String filepath=workingDir+"\\"+"Temp"+"\\"+CurrentDate+"."+"xlsx";
			 
			 resultfile = filepath;
			 String SheetName ="dbData";

			try {
				FileOutputStream fo = new FileOutputStream(filepath);
				
				XSSFWorkbook workbook = new XSSFWorkbook(); 
				
				 ws = workbook.createSheet(SheetName);
				 FormulaEvaluator wbevaluator = workbook.getCreationHelper().createFormulaEvaluator();
				if(update.equalsIgnoreCase("yes")){
				 wbevaluator.evaluateAll();
				}
			 
				
					if (!resultfile.endsWith(".xlsx")){
						System.out.println("Please correct Excel file extension, Data can be exported to only .xlsx file");
						LOGGER.log(Level.SEVERE, "Exception occured in executeQuery Please correct Excel file extension, Data can be exported to only .xlsx file");
						System.out.println("nofetchdata");
					}else{
				
					 rs.beforeFirst(); 
					 XSSFCell cell =null;
					 ResultSetMetaData rsmd= (ResultSetMetaData) rs.getMetaData();
					 
						int column = rsmd.getColumnCount();
					  
						String columnname = rsmd.getColumnLabel(1);
						 
						int rowcount=0;
						while(rs.next()){
							rowcount++;
						}
						 					
					 
						 		 
						 rs.beforeFirst();
					 	int tempj=0;
							while(rs.next()){                      //Writing Column Name in excel
								XSSFRow row = ws.createRow(rowcount);
								  
					     for(int c =0;c<column;c++){
					    cell = row.createCell(c);
					    	 
					    	  
										cell.setCellValue(rsmd.getColumnLabel(c+1)+0+" "+c);								
								 								
								}
					   
							}
					  
							rs.beforeFirst();
							int i=0;
							 tempj=1;
							while(rs.next()){
															
								for(int c =0;c<column;c++){
										 
										 
										cell.setCellValue(rs.getString(c+1)+tempj+" "+c);
																	 
									
								}
								
									
							tempj++;
							}
																	
							
							
							try {
								FileOutputStream fileOut = new FileOutputStream(filepath);
								
								workbook.write(fileOut);
								 
								workbook.close();
								fo.close();
								 
							 
							} catch (IOException e) {
								LOGGER.log(Level.SEVERE, "IO Exception Occured in executeQuery -  while Writing data into Excel"+e.getMessage());
								System.out.println("IO Exception Occured, while Writing data into Excel "+e);
								 
							}
							
					
					}	
					 
				} catch (SQLException e) {
					System.out.println(" SQLException occured in connection with DB, check your credentials are correct "+e);
					LOGGER.log(Level.SEVERE, " SQLException occured in executeQuery -  while connection with DB, check your credentials are correct "+e.getMessage());
				} catch (IOException e) {
					 System.out.println("IO exception occured in executeQuery function  "+e.getMessage());
					 LOGGER.log(Level.SEVERE, "IO Exception Occured in executeQuery -   "+e.getMessage());
				}
					}
			 	
	public void setBrowser(String BrowserType) {
		browsertypeflg =0;
		File folder  = new File(DriverPath);
		File[] listofFiles = folder.listFiles() ;
				
		try {
			System.out.println(BrowserType);

			switch (BrowserType.toUpperCase().trim()) {
			case "IE":
				browsertypeflg =0;
				for(int i=0;i<listofFiles.length;i++){
				 
					String filename = listofFiles[i].getName();
					if(filename.contains("Ie")||filename.contains("ie")||filename.contains("IE")){
						filename = listofFiles[i].toString();
						ieDriverPath =filename;
							
						break;
					}
				}
				
				System.setProperty("webdriver.ie.driver", ieDriverPath);
				// Changed on 16/12/2013
				DesiredCapabilities capability = DesiredCapabilities
						.internetExplorer();
				capability
						.setCapability(
								InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS,
								true);
				capability.setCapability("useLegacyInternalServer", true);
				capability.setCapability("nativeEvents", false);
				D8 = new InternetExplorerDriver(capability);
				D8.getWindowHandle();
				D8.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
				D8.manage().window().maximize();
				Capabilities cap = ((RemoteWebDriver) D8).getCapabilities();
				browserver = cap.getVersion();
				browser="IE";

				break;
			case "FF":
				
				browsertypeflg =0;
				D8 = new FirefoxDriver();
				D8.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
				D8.manage().window().maximize();
				browser="FF";
				break;
				
				/*browsertypeflg =0;
				for(int i=0;i<listofFiles.length;i++){
				 
					String filename = listofFiles[i].getName();
					if(filename.contains("gecko")||filename.contains("GECKO")||filename.contains("Gecko")){
						filename = listofFiles[i].toString();
						ffDriverPath =filename.toString();
							
						break;
					}
				}
				System.setProperty("webdriver.gecko.driver", ffDriverPath);
				
				DesiredCapabilities capabilities = DesiredCapabilities.firefox();
				capabilities.setCapability("marionette", true);
			  D8 = new FirefoxDriver(capabilities);
			
				//D8 = new FirefoxDriver();
				D8.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
				D8.manage().window().maximize();
				browser="FF";
				break;*/
				
			case "GC":
				for(int i=0;i<listofFiles.length;i++){
					 
					String filename = listofFiles[i].getName();
					if(filename.contains("chrome")||filename.contains("Chrome")||filename.contains("CHROME")){
						filename = listofFiles[i].toString();
						chromeDriverPath =filename.toString();
						 
						break;
					}else{
						System.out.println("No Valid Driver server is provided");
						 
					}
				}
				
		System.setProperty("webdriver.chrome.driver", chromeDriverPath);
		 ChromeOptions options = new ChromeOptions();
		 

				DesiredCapabilities gccapability = DesiredCapabilities.chrome();
			gccapability.setCapability("useLegacyInternalServer", true);
			gccapability.setCapability("nativeEvents", false);
			boolean js = gccapability.isJavascriptEnabled();
			//js=false;
				
			gccapability.setCapability(ChromeOptions.CAPABILITY, options);
	 		
				D8 = new ChromeDriver(gccapability);
 
			D8.getWindowHandle();
				D8.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
				D8.manage().window().maximize();
				Capabilities gccap = ((RemoteWebDriver) D8).getCapabilities();
			browserver = gccap.getVersion();
				browsertypeflg=1;
				browser="GC";
				break;

			default:
				LOGGER.log(Level.SEVERE, "Error : Invalid browser type");
				System.out.println("Error : Invalid browser type");
			}
		} catch (Exception e1) {
			LOGGER.log(Level.SEVERE, "Error :Check the Test Suite file contents"+e1.getMessage());
			System.out.println("Check the Test Suite file contents");
			System.out.println(e1);
		}
	}
	
	public void findExecTestscript(String TestSuite, String TestScript,
			String ObjectRepository) throws Exception {
		try{
		 		
			
			FileInputStream fs = null;
			
			fs = new FileInputStream(new File(TestSuite));
			
			XSSFWorkbook TSworkbook = new XSSFWorkbook(fs);
			XSSFSheet TSsheet = TSworkbook.getSheetAt(0);
	    	 boolean ValidTestSuit=false;
			
			try {
				fs = new FileInputStream(new File(TestSuite));
			} catch (Exception e) {
				System.out
						.println("No proper TestSuite is defined in Selenium Utility file. Give the correct TestSuite name");
				System.out.println(e);
				LOGGER.log(Level.SEVERE, "Error :No proper TestSuite is defined in Selenium Utility file. Give the correct TestSuite name"+e.getMessage());
				return;
			}
			 
		
			cur_dt = new Date();
			DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
			String strTimeStamp = dateFormat.format(cur_dt);
			String[] dateArray = strTimeStamp.split("-");
			today = Integer.parseInt(dateArray[2]);
			String rp = reportsPath;
			 
			if (rp == "") {
				// if results path is passed as null, by
				// default 'C:\' drive is used

				// Updated on on 12/16/2013
				rp = tempTestReportPath;
			}

			if (!rp.endsWith("\\")) { // checks whether the path ends with
										// '/'
				rp = rp + "\\";
			}
			 
			strResultPath = rp + "Log";
			System.out.println();
			 
			
			String htmlname1 = rp + "Log" + "Test_Suite_" +  strTimeStamp+"_"+ browser
					+ ".html";
			File f = new File(strResultPath);
			if (strResultPath != f.getAbsolutePath().toString()) {
				System.out
						.println("Report will be printed in the following path since THE REPORT PATH WAS NOT GIVEN / THE GIVEN PATH IS INCORRECT.");
				LOGGER.log(Level.SEVERE, "WARNING :Report will be printed in the following path since THE REPORT PATH WAS NOT GIVEN / THE GIVEN PATH IS INCORRECT." );
				System.out.println(f.getAbsolutePath().toString());
			}
			f.mkdirs();
			bw1 = new BufferedWriter(new FileWriter(htmlname1));
			bw1.write("<HTML><BODY><TABLE BORDER=0 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
			bw1.write("<TABLE BORDER=0 BGCOLOR=BLACK CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
			bw1.write("<TR><TD COLSPAN=6 BGCOLOR=#66699><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Testcase Name</B></FONT></TD><TD BGCOLOR=#66699 WIDTH=27%><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Status</B></FONT></TD></TR>");

	 
			Iterator<Row> iterator = TSsheet.iterator();
			  	
			
			  String run="r";
		        boolean runscript =false;
		        while (iterator.hasNext()) {
		            Row nextRow = iterator.next();
		            Iterator<Cell> cellIterator = nextRow.cellIterator();
		             
		            while (cellIterator.hasNext()) {
		                Cell cell = cellIterator.next();
		                 
		                switch (cell.getCellType()) {
		                    case Cell.CELL_TYPE_STRING:
		                            if(cell.getStringCellValue().equalsIgnoreCase(run)){
		                        	runscript =true;
		                        	scriptfound = 1;
		                          }
		                    
		                    	 
		                        if(runscript==true){
		                        	cell =cellIterator.next();
		                        	 		                        	 
		                        	        	runscript=false;
		                        	        	
		                        	 ScriptName = cell.getStringCellValue();
		                        	 System.out.println(" ScriptName  "+ScriptName);
		                        	
		                        	if(!ScriptName.endsWith(".xlsx")){
		                        		
		                        		System.out
										.println("Invalid File Format of TestScript: "
												+ ScriptName
												+ "in Test Suite.Give only xlsx format");
		                        		LOGGER.log(Level.SEVERE, "Invalid File Format of TestScript: "
												+ ScriptName
												+ "in Test Suite.Give only xlsx format" );
		                        	}
		                        	try{
		                        		
		                        		
		                        	execKeywordScript(ScriptName, TestScript,
		    								ObjectRepository);
		                        	}catch(Exception ex){
		                        		System.out.println("Function can not called   "+ex.getMessage());
		                        		LOGGER.log(Level.SEVERE, "Function can not called   "+ex.getMessage() );
		                        		
		                        		LOGGER.log(Level.SEVERE, "Invalid File: "
												+ ScriptName
												+ "is not available.Give the correct TestScript" );
		                        		
		                        		System.out
										.println("Invalid File:"
												+ ScriptName
												+ "   or  Runnable script is not available in given File. Give the correct TestScript");
		                        	}
		                                               
		                        	
		                        	if ((exeStatus.equalsIgnoreCase("Failed"))||(exeStatus_suite.equalsIgnoreCase("Failed"))) {                       		 
		        				
		        				
		        						/*strResultPath pointing towards Detail report*/	   
        								// 2016-03-28 A pass of a reference previous HTML was corrected in a relative path.
		                        		bw1.write("<TR><TD COLSPAN=6 BGCOLOR=WHITE><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>"
		        								+ "<a href=" + DetailReport + ">"+ TCNm[0] +"</a>"
		        								+ "</B></FONT></TD><TD BGCOLOR=WHITE WIDTH=27%><FONT FACE=VERDANA COLOR=RED SIZE=2><B>"
		        								+ exeStatus + "</B></FONT></TD></TR>");
		                        		exeStatus_suite = "pass";
		        					} else {
		        						 
		        					
		        						/*strResultPath pointing towards Detail report*/	   
		        						// 2016-03-28 Text color of a status(Pass) result of the　Summary report was corrected.
        								// 2016-03-28 A pass of a reference previous HTML was corrected in a relative path.
		        						bw1.write("<TR><TD COLSPAN=6 BGCOLOR=WHITE><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>"
		        								+ "<a href=" + DetailReport + ">"+ TCNm[0] +"</a>"
		        								+ "</B></FONT></TD><TD BGCOLOR=WHITE WIDTH=27%><FONT FACE=VERDANA COLOR=GREEN SIZE=2><B>"
		        								+ exeStatus + "</B></FONT></TD></TR>");
		        					}
		                        	
		                        }
		                        
		                        break;
		                    case Cell.CELL_TYPE_BOOLEAN:
		                        System.out.print(cell.getBooleanCellValue());
		                        break;
		                    case Cell.CELL_TYPE_NUMERIC:
		                        System.out.print(cell.getNumericCellValue());
		                        break;
		                        
		                    case Cell.CELL_TYPE_BLANK:        
		                    	 if(cell.getStringCellValue().equalsIgnoreCase(run)){
			                        	runscript =true;
			                        	scriptfound = 1;
			                          }
		                    	 break;
		                    case Cell.CELL_TYPE_FORMULA:        
		                    	 if(cell.getStringCellValue().equalsIgnoreCase(run)){
			                        	runscript =true;
			                        	scriptfound = 1;
			                          }
		                    	 break;
		                }
		              
		            }
		          
		        }
			  
		    	if (!(scriptfound == 1)) {
					System.out
							.println("Invalid File Format of Testsuite/ No Test Script to execute. Give the correct Testsuite/TestScript. ");
					LOGGER.log(Level.SEVERE, "Invalid File Format of Testsuite/ No Test Script to execute. Give the correct Testsuite/TestScript.  "  );
				}

				bw1.close();
		        TSworkbook.close();
		        fs.close();
		        
		}catch(Exception e1){
			System.out.println("Error Occured in FindExecTestscript function "+e1.getMessage());
			LOGGER.log(Level.SEVERE, "Error Occured in FindExecTestscript function  "+e1.getMessage() );
			 
			
			bw1.close();
		}
		}
	
	public void execKeywordScript(String scriptName, String TestScript,
			String ObjectRepository) throws Exception

	{
		 
		FileInputStream fs2 = null;
		int cCellDatanum =0;
		// Report header
		cur_dt = new Date();
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
		String strTimeStamp = dateFormat.format(cur_dt);

		if (reportsPath == "") {

			reportsPath = tempTestReportPath;
		}

		if (!reportsPath.endsWith("\\"))

		{

			reportsPath = reportsPath + "\\";
		}
		TCNm = scriptName.split("\\.");
		strResultPath = reportsPath + "Log" + "\\" + TCNm[0] + "\\";
		String htmlname = reportsPath + "Log" + "\\" + TCNm[0] + "\\"
				+ strTimeStamp + ".html";
 
		reportsFolder = reportsPath + "Log" + "\\" + TCNm[0] + "\\";
		

		// 2016-03-28 A pass of a reference previous HTML was corrected in a relative path.
		DetailReport = ".\\Log" + "\\" + TCNm[0] + "\\"
				+ strTimeStamp + ".html";
		
		if (!(screenReport.contains("screen"))) {
			screenReport = screenReport + "Screenshot\\";

		}
		screenReport = screenReport.substring(0, screenReport.length()-1);
		 
		ScreenshotPath=screenReport+"\\"+TCNm[0];
		
 
	
		
		File f = new File(strResultPath);
		f.mkdirs();
		bw = new BufferedWriter(new FileWriter(htmlname));
		bw.write("<HTML><BODY><TABLE BORDER=0 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
		bw.write("<TABLE BORDER=0 BGCOLOR=BLACK CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
		
		bw.write("<TR><TD BGCOLOR=#66699 WIDTH=27%>"
				+ "<FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>Test Case Name:</B></FONT></TD>"
				+ "<TD COLSPAN=6 BGCOLOR=#66699><FONT FACE=VERDANA COLOR=WHITE SIZE=2><B>"
				+"<a href=\"file:///"+ScreenshotPath+"\">"+ TCNm[0] + "</a>"+"</B></FONT></TD></TR>");
	 
		
		
		
		bw.write("<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>");
		
		
		bw.write("<TR COLS=7><TD BGCOLOR=#FFCC99 WIDTH=3%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Row</B></FONT></TD>"
				+ "<TD BGCOLOR=#FFCC99  WIDTH=3%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Keyword</B></FONT></TD>"
				+ "<TD BGCOLOR=#FFCC99 WIDTH=20% ><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Object</B></FONT></TD>"
				+ "<TD BGCOLOR=#FFCC99  WIDTH=23%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Action</B></FONT></TD>"
				+ "<TD BGCOLOR=#FFCC99 WIDTH=21%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Comments</B></FONT></TD>"
				+ "<TD BGCOLOR=#FFCC99 WIDTH=25%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Execution Time</B></FONT></TD>"
				+ "<TD BGCOLOR=#FFCC99 WIDTH=5%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Status</B></FONT></TD>");
		
		

		exeStatus = "Pass";
		exeStatus_suite="Pass";
		
		if (!TestScript.endsWith("\\"))

		{

			TestScript = TestScript + "\\";
		}
		
		String scriptPath = TestScript + scriptName;
		 
	 
		TScname = scriptName;
	 
		FileInputStream fs1 = null;
		 
		fs1 = new FileInputStream(new File(scriptPath));
	
		XSSFWorkbook TScworkbook = new XSSFWorkbook(fs1);
		TScsheet = TScworkbook.getSheetAt(0);
		
		FormulaEvaluator evaluator = TScworkbook.getCreationHelper().createFormulaEvaluator();
		
		
		TScrowcount = TScsheet.getLastRowNum()+1;
		 rowcnt = 0;

		for (j = 0; j < TScrowcount; j++)

		{
		 
			rowcnt = rowcnt + 1;
			String TSvalidate = "r";
			
			if((TScsheet.getRow(j).getCell(0).getStringCellValue().equalsIgnoreCase(TSvalidate)==true))	{
				                                                                                                                                                                          
		try{
				Action =TScsheet.getRow(j).getCell(1).getStringCellValue();
			
				
				 if(TScsheet.getRow(j).getCell(2).getCellType()== XSSFCell.CELL_TYPE_STRING)   {
					 cCellData = TScsheet.getRow(j).getCell(2).getStringCellValue();  
				 }
				 else     if(TScsheet.getRow(j).getCell(2).getCellType()== XSSFCell.CELL_TYPE_NUMERIC)  {                                                                                                                                  
				 
					 waitcCellDatanum =(int) TScsheet.getRow(j).getCell(2).getNumericCellValue();
					 cCellData = Integer.toString(waitcCellDatanum);
				 
			 
				  }
				 else     if(TScsheet.getRow(j).getCell(2).getCellType()== XSSFCell.CELL_TYPE_FORMULA)  {                                                                                                                                  
					 
					 evaluator.evaluateFormulaCell(TScsheet.getRow(j).getCell(2));
					 cCellData=TScsheet.getRow(j).getCell(2).getStringCellValue();  
					  }
				 else     if(TScsheet.getRow(j).getCell(2).getCellType()== XSSFCell.CELL_TYPE_BLANK)  {                                                                                                                                  
					 cCellDatanum = (int) TScsheet.getRow(j).getCell(2).getNumericCellValue();  
					// cCellData="";
					  }
				 
				 
				  if(TScsheet.getRow(j).getCell(3).getCellType()==XSSFCell.CELL_TYPE_BLANK){
					 dCellData =TScsheet.getRow(j).getCell(3).getStringCellValue();
				 }else if(TScsheet.getRow(j).getCell(3).getCellType() == XSSFCell.CELL_TYPE_BLANK){
					 dCellData="";
				 }
				  else 	 if(TScsheet.getRow(j).getCell(3).getCellType()==XSSFCell.CELL_TYPE_FORMULA){
					  
				 if(update.equalsIgnoreCase("yes")){
			 			 evaluator.evaluateAll();
					  dCellData = TScsheet.getRow(j).getCell(3).getStringCellValue();
				}
				else {
			 				dCellData = TScsheet.getRow(j).getCell(3).getStringCellValue();
					}
				
				 }
				  else if(TScsheet.getRow(j).getCell(3).getCellType()==XSSFCell.CELL_TYPE_STRING){
					 dCellData = TScsheet.getRow(j).getCell(3).getStringCellValue();
				 } 
				  if(TScsheet.getRow(j).getCell(3).getCellType()==XSSFCell.CELL_TYPE_NUMERIC){
						 double  dCellData1 = TScsheet.getRow(j).getCell(3).getNumericCellValue();
					  	  dCellData =String.valueOf(dCellData1);
					 }
				  
				  				  
				if(TScsheet.getRow(j).getCell(5).getCellType() == XSSFCell.CELL_TYPE_BLANK){
					sComments ="";
					comments =sComments;
				}else if(TScsheet.getRow(j).getCell(5).getCellType() == XSSFCell.CELL_TYPE_NUMERIC){
					nComments = TScsheet.getRow(j).getCell(5).getNumericCellValue();
					comments = Double.toString(nComments);
				}else if(TScsheet.getRow(j).getCell(5).getCellType()==XSSFCell.CELL_TYPE_STRING){
					sComments = TScsheet.getRow(j).getCell(5).getStringCellValue();
					comments = sComments;
				}
		}catch(Exception e){
			System.out.println("Error Occured, while reading TestScript File "+e.getMessage());
			LOGGER.log(Level.SEVERE, "Error Occured, while reading TestScript File in execKeywordScript  "+e.getMessage() );
		}
				String ORPath = ObjectRepository;
				
			
				try {
					fs2 = new FileInputStream(new File(ORPath));
					XSSFWorkbook ORworkbook = new  XSSFWorkbook(fs2);
					
					ORsheet = ORworkbook.getSheetAt(0);
					
					ORrowcount =ORsheet.getLastRowNum();
					
					 
					ActionVal = Action.toLowerCase();
					iflag = 0;

				} catch (Exception e)
				{
					 
					LOGGER.log(Level.SEVERE, "Error Occured  in execKeywordScript  "+e.getMessage() );
					System.out.print(e.getMessage());
					fail("Object Repository format is not correct.Check the File format");
					
				}
				
				bCellAction(scriptName);
				System.out.println((j+1)+" | "+Action+" | "+cCellData+" | "+dCellData+" | "+comments);
				
				
			}

		}
		bw.close();
		fs1.close();
		fs2.close();

	
	}

	public static void screenShot(int loopn, int rown, String Sname)
			throws Exception {
		try {
			if (devFlag == 0) {
				screenshotflag = screenshotflag + 1;
				DateFormat dateFormat = new SimpleDateFormat(
						"yyyyMMdd-HH-mm-ss");
				Date date = new Date();
				String filenamer = "";
				String strTime = dateFormat.format(date);
				Sname = Sname.substring(0, Sname.indexOf("."));
				ScreenFolder = Sname;
			 
				ScreenshotPath = screenReport + Sname;
 

				screenReport = screenReport.toLowerCase();
				if (screenReport == "") {
					screenReport = reportsPath;
					initscreenFlag = 2;
				}

				if (!screenReport.endsWith("\\")) {
					screenReport = screenReport + "\\";

				}
				if (!(screenReport.contains("screen"))) {
					screenReport = screenReport + "Screenshot\\"+browser+"\\";

				}
				initscreen = initscreen + 1;
				//when path is not given use temp folder "TempTestReportPath"
				if (initscreen == 2 && initscreenFlag == 2) {
					System.out
							.println("Screenshots will be captured in the following path since the SCREENSHOTS PATH IS NOT GIVEN.");
					LOGGER.log(Level.SEVERE, "Screenshots will be captured in the following path since the SCREENSHOTS PATH IS NOT GIVEN    "  );
					System.out.println("screenReport (initscreen == 2 && initscreenFlag == 2) ==>"+screenReport);
					
				}

				else {
					if (initscreen == 2) {
						File f1s = new File(screenReport);
					 
						if (!screenReport.substring(0, screenReport.length() - 1).equalsIgnoreCase(f1s.getAbsolutePath().toString()))

						{
							System.out
									.println("Screenshots will be captured in the following path since the given SCREENSHOTS PATH IS NOT AVAILABLE.");
							LOGGER.log(Level.SEVERE, "Screenshots will be captured in the following path since the given SCREENSHOTS PATH IS NOT AVAILABLE   " +f1s.getAbsolutePath().toString() );
							System.out
									.println(" initscreen == 2"+f1s.getAbsolutePath().toString());

						}
						
					}
				}
				if (loopflag == 0) {

					filenamer = screenReport + Sname + "\\" + Sname + "_"+browser+"_"
							+ screenshotflag + "_rowno_" + (j + 1) + "_"
							+ strTime + ".png";

				} else {

					filenamer = screenReport + Sname + "\\" + Sname + "_"+browser+"_"
							+ screenshotflag + "_loop_"
							+ (loopcnt[loopsize] + 1) + "_rowno_" + (j + 1)
							+ "_" + strTime + ".png";
				}
	
		 
		 String folderlocation =screenReport +Sname+ "\\";
 
		 new File(folderlocation).mkdirs();

				if ((ScreenshotTypeFlag == 0)&&((browsertypeflg!=1))){						
				
					final Screenshot screenshot = new AShot().takeScreenshot(D8);
				 
				       final BufferedImage image = screenshot.getImage();
				        ImageIO.write(image, "PNG", new File(filenamer));
		
				} else{
				  if((ScreenshotTypeFlag == 1)) {
				
										
					BufferedImage image = new Robot()
							.createScreenCapture(new Rectangle(Toolkit
									.getDefaultToolkit().getScreenSize()));
					File outputfile = new File(filenamer);
					//outputfile.mkdirs();
					ImageIO.write(image, "png", outputfile);
					Thread.sleep(3000);
				}}
		
				if(browsertypeflg==1){
							
					 final Screenshot screenshot = new AShot().shootingStrategy(ShootingStrategies.viewportPasting(500)).takeScreenshot(D8);
				     final BufferedImage image = screenshot.getImage(); 
				         
				        ImageIO.write(image, "PNG", new File(filenamer));
				        
				}
			}
			/*
			 * else { Update_Report("notForDevice1"); }"
			 */
		}

		catch (Exception e) {
			System.out
					.println("Screenshot was not printed. Check the file path");
			System.out.println("Actual error "+e.getMessage());
			 e.printStackTrace();
			LOGGER.log(Level.SEVERE, "Screenshot was not printed. Check the file path  "  );
			update_Report("screenshot", e);
		}
	}

	public static void update_Report(String Res_type) throws IOException {
		try {
			
			String str_time;
		
			Date exec_time = new Date();
			DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
			str_time = dateFormat.format(exec_time);
		 
			if (Res_type.startsWith("executed")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = GREEN>"
						+ "Passed" + "</FONT></TD></TR>");
			} else if (Res_type.startsWith("failed")) {
				exeStatus = "Failed";
				exeStatus_suite="Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");

			} else if (Res_type.startsWith("NoWindowFound")) {
				exeStatus = "Failed";
				exeStatus_suite="Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Fail" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number: "
						+ (j + 1)
						+ ". Description: The Window not Found</div></th></FONT></TR>");
			} else if (Res_type.startsWith("loop")) {

				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=BLUE><div align=left></FONT><FONT FACE=VERDANA SIZE=2 COLOR = BLUE>"
						+ Res_type + "</div></th></FONT></TR>");
			} else if (Res_type.startsWith("callactionstart")) {

				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=BLUE><div align=left></FONT><FONT FACE=VERDANA SIZE=2 COLOR = BLUE>"
						+ " Start of CallAction : '"
						+ globalCompName
						+ "' execution" + "</div></th></FONT></TR>");
			} else if (Res_type.startsWith("callactionend")) {

				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=BLUE><div align=left></FONT><FONT FACE=VERDANA SIZE=2 COLOR = BLUE>"
						+ " End of CallAction : '"
						+ globalCompName
						+ "' execution" + "</div></th></FONT></TR>");
			} else if (Res_type.startsWith("callactionfnf")) {
				exeStatus = "Failed";
				exeStatus_suite="Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Fail" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number: "
						+ (j + 1)
						+ ". Description: The Action(TestScript) name given is not available in the given path. Check the file path and action name.</div></th></FONT></TR>");
			} else if (Res_type.startsWith("callactionff")) {
				exeStatus = "Failed";
				exeStatus_suite="Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Fail" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number: "
						+ (j + 1)
						+ ". Description: The Action(TestScript) format is not in supported. Give the file name with xlsx extension only.</div></th></FONT></TR>");
			}

			else if (Res_type.startsWith("missing")) {
				exeStatus = "Failed";
				exeStatus_suite="Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = ORANGE>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number "
						+ (j + 1)
						+ ".Description: The Datatable column name not found</div></th></FONT></TR>");
			} else if (Res_type.startsWith("ObjectLocator")) {
				exeStatus = "Failed";
				exeStatus_suite="Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = ORANGE>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number "
						+ (j + 1)
						+ ". Description: Object Locator is wrong or not supported. Supported Locators are Id,Name,Xpath& CSS</div></th></FONT></TR>");
			}

			else if (Res_type.startsWith("tooManyArguments")) {
				exeStatus = "Failed";
				exeStatus_suite="Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED >"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "WARNING" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>WARNING @ Line NO:  "
						+ (j + 1)
						+ ". Pass only one argument if you are using data or environment variables. Otherwise only the first variable will be considered</div></th></FONT></TR>");
			} else if (Res_type.startsWith("NoBlankAvailable")) {
				exeStatus = "Failed";
				exeStatus_suite="Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED >"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error @ Line NO:  "
						+ (j + 1)
						+ ". No Blank Value is available in the Element</div></th></FONT></TR>");
			} else if (Res_type.startsWith("objNotFound")) {
				exeStatus = "Failed";
				exeStatus_suite="Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in test step number "
						+ (j + 1)
						+ ". Object with the given object name is not found in objectrepository</div></th></FONT></TR>");

			} else if (Res_type.startsWith("keyword")) {
				exeStatus = "Failed";
				exeStatus_suite="Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ Action
						+ ": Keyword not supported by Open2Test</div></th></FONT></TR>");
			} else if (Res_type.startsWith("nodatatable")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Data table not available, Ensure whether you have imported the right Data Table.</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("action")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ dCellData
						+ " : Action not supported by Open2Test</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("action1")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ dCellData
						+ " : The action is not supported for this type of object</div></th></FONT></TR>");
			} else if (Res_type.startsWith("objectmiss")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number "
						+ (j + 1)
						+ ".Object not found in the page</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("property")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Property not supported by Open2Test</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("property1")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Property not supported for this kind of object</div></th></FONT></TR>");
			} else if (Res_type.startsWith("condFailed")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (conditionline + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = ORANGE>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Condition Failed. Statements within the condition and End Condition will not be executed</div></th></FONT></TR>");

			} else if (Res_type.equalsIgnoreCase("invaliddate1")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";

				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Invalid date has been given. Give the proper date format(It should be mm-dd-yyyy format) for a valid date in test script.</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("invaliddate")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Given Date cannot be set. Check the date is valid and in allowed range</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("prevmonth")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Element releated to the previous month is not identified in the page. Try with different idenfication class by adding the same in setting.java.</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("nextmonth")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Element releated to the next month is not identified in the page. Try with different idenfication class by adding the same in setting.java.</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("titlemonth")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Element releated to the title month is not identified in the page. Try with different idenfication class by adding the same in setting.java.</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("titleyear")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Element releated to the titleyear is not identified in the page. Try with different idenfication class by adding the same in setting.java.</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("titledefault")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Element releated to the title year and month  is not identified in the page. Try with different idenfication class by adding the same in setting.java.</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("monthnotidentified")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Month in the calendar is not identified. English and Japanese character set are supported.</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("invalidmonth")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";

				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Given month is invalid. Correct the month and retry.</div></th></FONT></TR>");

			} else if (Res_type.equalsIgnoreCase("filePathNotFound")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";

				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Given file does not exist.Check the File path and File Name.</div></th></FONT></TR>");

			} else if (Res_type.equalsIgnoreCase("filePathNotFound1")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";

				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Given file path does not exist.Check the File path.</div></th></FONT></TR>");

			} else if (Res_type.equalsIgnoreCase("filePathNotFound2")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";

				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "FilePath is not Given.Give a valid file path</div></th></FONT></TR>");

			} else if (Res_type.equalsIgnoreCase("calendaraction")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";

				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "For SetDate Operation ObjectType Should be calendar and Object Name should start with cal_</div></th></FONT></TR>");
			} else if (Res_type.equalsIgnoreCase("userdefined")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Function is not executed. Error exists in  User Defined function. Correct the User Defined Function</div></th></FONT></TR>");
			}

			else if (Res_type.equalsIgnoreCase("getcelldata")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Invalid Syntax for getcelldata. Pass Exact number of arguments</div></th></FONT></TR>");
			}
			else if (Res_type.equalsIgnoreCase("nofetchdata")) {
				exeStatus_suite="Failed";
				 
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Invalid Syntax for Excelfile.  Data can be exported to only .xlsx file</div></th></FONT></TR>");
			}else if (Res_type.equalsIgnoreCase("NoColumnFound")) {
				exeStatus_suite="Failed";
				 
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "NO Column found in Data Sheet-which mentioned as dt_XXX; Please verify your Data sheet and column</div></th></FONT></TR>");
			}else if (Res_type.equalsIgnoreCase("NoMatchinDataTable")) {
			 
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "NO Matching text found in Data Sheet- for table element which you mentioned in script; </div></th></FONT></TR>");
			}
			else if (Res_type.equalsIgnoreCase("ObjectNotFound")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Object details you mentioned in c column of script file, dosen't match with the Object repository </div></th></FONT></TR>");
			}else if (Res_type.equalsIgnoreCase("invalidQuery")) {
				
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Given Query is not Valid , Please enter correct Select query </div></th></FONT></TR>");
			}else if (Res_type.equalsIgnoreCase("invalidConnection")) {
				
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number. "
						+ (j + 1)
						+ ". "
						+ "Not able to Connect to Given Data base, please check the coneection parameters or appropriate Jar files are uploaded </div></th></FONT></TR>");
			}

		} catch (Exception e1) {
			System.out
					.println("Report will not be printed. Check the file path.  "
							+ e1);
		}

	}

	public static void update_Report(String Res_type, Exception e)
			throws IOException {
		try {
			String str_time;
			Date exec_time = new Date();
			DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
			str_time = dateFormat.format(exec_time);

			exeStatus = "Failed";
			if (Res_type.startsWith("failed")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left></FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ e.toString().substring(
								e.toString().indexOf(":") + 1,
								e.toString().indexOf(".",
										e.toString().indexOf(":") + 1) + 1)
						+ "</div></th></FONT></TR>");
			} else if (Res_type.startsWith("keyword")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number "
						+ (j + 1)
						+ ". Description: Keyword not supported by Open2Test</div></th></FONT></TR>");
			} else if (Res_type.startsWith("screenshot")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"

						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ dCellData
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in  test step number "
						+ (j + 1)
						+ ". Screenshot not printed</div></th></FONT></TR>");
			}
		} catch (Exception e4) {
			System.out
					.println("Report will not be printed. Check the file path."
							+ e4);
		}

	}
		
	public static void update_Report(String Res_type, String ObjectSet,
			String ObjectSetVal) throws IOException // //Updated on 16.12.2013
													// for Adding the values
													// instead variable names in
													// the report
	{
		try {
			String str_time;
			
			Date exec_time = new Date();
			DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
			str_time = dateFormat.format(exec_time);
			if (reporttype == 1) {
			} else {
			}
			if (Res_type.startsWith("executed")
					&& ObjectSet.equalsIgnoreCase("page")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ ObjectSet
						+ " ; "
						+ ObjectSetVal
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = GREEN>"
						+ "Passed" + "</FONT></TD></TR>");
			} else if (Res_type.startsWith("executed")) {
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ ObjectSet
						+ " : "
						+ ObjectSetVal
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = GREEN>"
						+ "Passed" + "</FONT></TD></TR>");
			} else if (Res_type.startsWith("NoWindowFound")
					&& ObjectSet.equalsIgnoreCase("page")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ ObjectSet
						+ " ; "
						+ ObjectSetVal
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number: "
						+ (j + 1)
						+ ". Description: The Window not Found</div></th></FONT></TR>");
			}

			else if (Res_type.startsWith("failed")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ ObjectSet
						+ " : "
						+ ObjectSetVal
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
			} else if (Res_type.startsWith("loop")) {

				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=BLUE><div align=left></FONT><FONT FACE=VERDANA SIZE=2 COLOR = BLUE>"
						+ Res_type + "</div></th></FONT></TR>");
			} else if (Res_type.startsWith("missing")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ ObjectSet
						+ " : "
						+ ObjectSetVal
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = ORANGE>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number "
						+ (j + 1)
						+ ". Description: The Datatable column name not found</div></th></FONT></TR>");
			} else if (Res_type.startsWith("ObjectLocator")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ ObjectSet
						+ " : "
						+ ObjectSetVal
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = ORANGE>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in Keyword test step number "
						+ (j + 1)
						+ ". Description: Object Locator is wrong or not supported. Supported Locators are Id,Name,Xpath& CSS</div></th></FONT></TR>");
			} else if (Res_type.startsWith("Property")) {
				exeStatus_suite="Failed";
				exeStatus = "Failed";
				bw.write("<TR COLS=6><TD BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE=VERDANA SIZE=2>"
						+ (j + 1)
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=17%><FONT FACE=VERDANA SIZE=2>"
						+ Action
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ cCellData
						+ "</FONT></TD></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ ObjectSet
						+ " : "
						+ ObjectSetVal
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ comments
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>"
						+ str_time
						+ "</FONT></TD><TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2 COLOR = RED>"
						+ "Failed" + "</FONT></TD></TR>");
				bw.write("<TR COLS=6><th colspan= 6 BGCOLOR=#EEEEEE WIDTH=5%><FONT FACE='WINGDINGS 2' SIZE=3 COLOR=RED><div align=left>X </FONT><FONT FACE=VERDANA SIZE=2 COLOR = RED>Error Occurred in test step number "
						+ (j + 1)
						+ ". This property is not supported by Open2Test</div></th></FONT></TR>");
			}

		} catch (Exception e1) {
			System.out
					.println("Report will not be printed. Check the file path."
							+ e1);
		}
	}
	
	private static void func_StoreCheck() throws Exception {

		 

		try {
			String actval = null;
			String expval = null;
			Boolean boolval = null;
			String varname;
			String[] cCellDataValCh = cCellData.split(";");
			String objectType = cCellDataValCh[0];     
			String ObjectValCh = "";
			try {
				// Inserted on 16/12/2012. For holding Object Type
				ObjectValCh = cCellDataValCh[1];                      

				// Inserted on 16/12/2012. For Holding Object Name
			} catch (ArrayIndexOutOfBoundsException e) {
				ObjectValCh = "";
			}
			String[] dCellDataValCh = dCellData.split(":");

			String ObjectSetCh = dCellDataValCh[0];                             

			String ObjectSetValCh = "";
			int DTcolumncountCh = 0;
			if (objectType.equalsIgnoreCase("textelement")) {
				ObjectSetValCh = dCellData.replaceFirst("text:", "");
				  System.out.println("ObjectSetValCh   is ><> > "+ObjectSetValCh);
			} else {

				if (dCellDataValCh.length == 2) {
					ObjectSetValCh = dCellDataValCh[1];                    
				}
				// 2016-03-23　Multiple Selected Check Support
				else if (dCellDataValCh.length > 2) {
					ObjectSetValCh = dCellDataValCh[1];
					for(int i = 2; i < dCellDataValCh.length; i++) {
						ObjectSetValCh += ":" + dCellDataValCh[i];
					}
				}
			}
			if (dCellDataValCh.length == 2) {
				if (ObjectSetValCh.substring(0, 1).equalsIgnoreCase("#")) {
					System.out.println("In #");
					if (map.get(ObjectSetValCh.substring(1,
							(ObjectSetValCh.length()))) != null) {
						ObjectSetValCh = map.get(ObjectSetValCh.substring(1,
								(ObjectSetValCh.length())));

					} else {
						ObjectSetValCh = "";
					}
				}
			}
			if (objectType.equalsIgnoreCase("page")
					|| objectType.equalsIgnoreCase("dialog")) {

				objFoundFlag = 1;

			} else

			{
				
				
				for (int k = 0; k < ORrowcount; k++) {
				

					
					if(((ORsheet.getRow(k).getCell(1).getStringCellValue()).equalsIgnoreCase(ObjectValCh)==true))
					
				                                                                                                          
					{
						String[] ORcellData = new String[3];
						
						ORcellData = ORsheet.getRow(k).getCell(4).getStringCellValue().split("=",2);
						
						
						

						ORvalname = ORcellData[0];
						ORvalue = ORcellData[1];
						k = ORrowcount + 1;
						objFoundFlag = 1;
						
					}

				}
			}

			if (objFoundFlag == 1) {
				 
				objFoundFlag = 0;
				if (dCellDataValCh.length == 2) {
					if (ObjectSetValCh.contains("dt_")) {
						iflag = 0;
						String ObjectSetValtableheader[] = ObjectSetValCh
								.split("_");
						int column = 0;
						String Searchtext = ObjectSetValtableheader[1];
						try {
							DTcolumncountCh = DTsheet.getRow(0).getLastCellNum();
							 
						} catch (NullPointerException e) {
							return;
						}

						for (column = 0; column < DTcolumncountCh; column++) {
							
							if(Searchtext.equalsIgnoreCase(DTsheet.getRow(0).getCell(column).getStringCellValue())==true)
					 						
							{
								ObjectSetValCh = DTsheet.getRow(dtrownum).getCell(column).getStringCellValue();
								
								
								iflag = 1;
								if (ObjectSetValCh.length() == 0) {
									ObjectSetValCh = "";
								}  
							}

						}
						if (iflag == 0) {
							ObjectSetValCh = "nodatarow";
						}
					}

					if (ObjectSetValCh.contains("dt_")) {
						String ObjectSetValtableheader[] = ObjectSetValCh
								.split("_");
						int column = 0;
						String Searchtext = ObjectSetValtableheader[1];
						DTcolumncountCh =DTsheet.getRow(0).getLastCellNum();
										

						for (column = 0; column < DTcolumncountCh; column++) {
						
							if(Searchtext.equalsIgnoreCase(DTsheet.getRow(0).getCell(column).getStringCellValue())==true)
													
							{
								
								ObjectSetValCh = DTsheet.getRow(dtrownum).getCell(column).getStringCellValue();
							

								iflag = 1;
								if (ObjectSetValCh.length() == 0) {
									ObjectSetValCh = "";
								}
							}

						}
						if (iflag == 0) {
							ORvalname = "exit";
						}
					}
				}
				switch (ObjectSetCh.toLowerCase()) {
				case "enabled":
					if (objectType.equalsIgnoreCase("textbox")
							|| objectType.equalsIgnoreCase("combobox")
							|| objectType.equalsIgnoreCase("listbox")
							|| objectType.equalsIgnoreCase("radiobutton")
							|| objectType.equalsIgnoreCase("button")
							|| objectType.equalsIgnoreCase("checkbox")
							|| objectType.equalsIgnoreCase("textarea")
							|| objectType.equalsIgnoreCase("image")
							|| objectType.equalsIgnoreCase("table")
							|| objectType.equalsIgnoreCase("link")
							|| objectType.equalsIgnoreCase("element")) {
						func_FindObj(ORvalname, ORvalue);
						boolval = elem.isEnabled();
						actval = boolval.toString();
					} else {
						update_Report("property1");
					}
					break;
				case "text":
					// Updated on 10.9.2015
					// Specifications change for STH 
					if (objectType.equalsIgnoreCase("button")) {
						func_FindObj(ORvalname, ORvalue);
						if (elem.getTagName().equalsIgnoreCase("button")) {
							actval = elem.getText();
						} else if (elem.getTagName().equalsIgnoreCase("input")) {
							actval = elem.getAttribute("value");
						} else {
							update_Report("property1");
						}
					} else if (objectType.equalsIgnoreCase("textbox")
							|| objectType.equalsIgnoreCase("textarea")) {

						func_FindObj(ORvalname, ORvalue);
						actval = elem.getAttribute("value");

					} else if (objectType.equalsIgnoreCase("textelement")
							|| objectType.equalsIgnoreCase("element")) {

						func_FindObj(ORvalname, ORvalue);
						actval = elem.getText();

					} else if (objectType.equalsIgnoreCase("combobox")
							|| objectType.equalsIgnoreCase("listbox")
							) {
						func_FindObj(ORvalname, ORvalue);
						// 2016-03-23 Multiple Selected Check Support
						List<WebElement> selectedList = new Select(elem).getAllSelectedOptions();
						actval = selectedList.get(0).getText();
						if(selectedList.size() > 1) {
							for(int i = 1; i < selectedList.size(); i++) {
								actval += ":" + selectedList.get(i).getText();
							}
						}
					
					} else {

						update_Report("property1");
					}
					break;
				case "value":
					// Updated on 10.9.2015
					// Specifications change for STH
					// 2016-03-23 Multiple Selected Check Support
					if(objectType.equalsIgnoreCase("checkbox")
							|| objectType.equalsIgnoreCase("combobox")
							|| objectType.equalsIgnoreCase("radiobutton")) {
						func_FindObj(ORvalname, ORvalue);
						actval = elem.getAttribute("value");
					} else if(objectType.equalsIgnoreCase("listbox")) {
						func_FindObj(ORvalname, ORvalue);
						List<WebElement> selectedList = new Select(elem).getAllSelectedOptions();
						actval = selectedList.get(0).getAttribute("value");
						if(selectedList.size() > 1) {
							for (int i = 1; i < selectedList.size(); i++) {
								actval += ":" + selectedList.get(i).getAttribute("value");
							}
						}
					} else {
						update_Report("property1");
					}
					break;
				case "visible":
					if (objectType.equalsIgnoreCase("textbox")
							|| objectType.equalsIgnoreCase("combobox")
							|| objectType.equalsIgnoreCase("listbox")
							|| objectType.equalsIgnoreCase("radiobutton")
							|| objectType.equalsIgnoreCase("button")
							|| objectType.equalsIgnoreCase("checkbox")
							|| objectType.equalsIgnoreCase("textarea")
							|| objectType.equalsIgnoreCase("image")
							|| objectType.equalsIgnoreCase("table")
							|| objectType.equalsIgnoreCase("link")
							|| objectType.equalsIgnoreCase("element")) {
						func_FindObj(ORvalname, ORvalue);
						boolval = elem.isDisplayed();
						actval = boolval.toString();
					} else {
						update_Report("property1");
					}

					break;
				case "checked":
					if ((objectType.equalsIgnoreCase("radiobutton")
							|| objectType.equalsIgnoreCase("checkbox") || objectType
								.equalsIgnoreCase("element"))) {
						func_FindObj(ORvalname, ORvalue);
						boolval = elem.isSelected();
						actval = boolval.toString();

					} else {
						update_Report("property1");
					}
					break;
				case "linktext":
					if (objectType.equalsIgnoreCase("link")) {
						func_FindObj(ORvalname, ORvalue);

						actval = elem.getText();
					} else {
						update_Report("Property1");
					}
					break;
				case "pagetitle":
					if (ObjectValCh != "") {
						actval = ObjectValCh;
					} else {
						actval = D8.getTitle();
					}
					break;

				case "exist":
					boolval = false;
					actval = boolval.toString();

					if ((objectType.equalsIgnoreCase("page")) == true

							&& (D8.getTitle().toString()
									.equalsIgnoreCase(ObjectValCh)) == true) {

						actval = "true";
					} else {
						if (objectType.equalsIgnoreCase("page")) {

							String currentWindowHandle = D8.getWindowHandle();
							int windowFound = 0;
							WebDriver newWindow = null;
							Set<String> al = new HashSet<String>();
							al = D8.getWindowHandles();
							Iterator<String> windowIterator = al.iterator();

							if (D8.getTitle().toString()
									.equalsIgnoreCase(ObjectValCh) != true) {
								while (windowIterator.hasNext()) {
									String windowHandle = windowIterator.next();
									newWindow = D8.switchTo().window(
											windowHandle);
									if (newWindow.getTitle().toString()
											.equalsIgnoreCase(ObjectValCh) == true) {
										boolval = true;
										actval = boolval.toString();
										windowFound = 1;
										break;
									}

								}
								if (windowFound != 1) {
									boolval = false;

									actval = boolval.toString();
								}
								D8.switchTo().window(currentWindowHandle);
							}

						} else {

							if (objectType.equalsIgnoreCase("dialog") == true) {
								try {

									Alert dialogExist = null;
									dialogExist = D8.switchTo().alert();
									if (dialogExist.toString() != null) {
										boolval = true;
										actval = boolval.toString();
									} else {
										boolval = false;
										actval = boolval.toString();
									}

								} catch (NoAlertPresentException e) {

									boolval = false;
									actval = boolval.toString();

								}

							}
						}

					}

					break;
				case "rowcount":
					func_FindObj(ORvalname, ORvalue);
					List<WebElement> rows = elem.findElements(By.tagName("tr"));
					Integer rowCount = rows.size();
					if (rowCount == 0) {
						
						rowCount = 1;
					}
					actval = rowCount.toString();
					break;

				case "columncount":
					WebElement headerRow = null;
					func_FindObj(ORvalname, ORvalue);
					List<WebElement> rows1 = elem
							.findElements(By.tagName("tr"));
					try {
						headerRow = rows1.get(rows1.size()-2);
					} catch (Exception Ed) {
						try
						{
							headerRow = rows1.get(1);
						}
						catch(Exception Ed1)
						{
							headerRow = rows1.get(0);
						}
					}

					List<WebElement> columns = headerRow.findElements(By
							.tagName("th"));
					Integer colCount = columns.size();
					if (colCount == 0) {
						List<WebElement> columns0 = headerRow.findElements(By
								.tagName("td"));
						colCount = columns0.size();
						if (colCount == 0) {
							WebElement columns1 = headerRow.findElement(By
									.tagName("th"));
							if (columns1 != null) {
								colCount = 1;
							} else {
								WebElement columns2 = headerRow.findElement(By
										.tagName("td"));
								if (columns2 != null) {
									colCount = 1;
								}
							}
						}

					}
					actval = colCount.toString();
					break;
				case "getcelldata":

					func_FindObj(ORvalname, ORvalue);
					try {
						List<WebElement> cellRows = elem.findElements(By
								.tagName("tr"));
						getCellArray = dCellData.split(":");
						String cellData = "";
						int rowNumber = Integer.parseInt(getCellArray[2]);
						int colNumber = Integer.parseInt(getCellArray[3]);
						System.out.println("rownum - " + rowNumber
								+ "  col num - " + colNumber);
						WebElement reqrow = cellRows.get(rowNumber - 1);
						List<WebElement> reqcolmns = reqrow.findElements(By
								.tagName("td"));
						WebElement reqcellData = reqcolmns.get(colNumber - 1);
						cellData = reqcellData.getText();
						actval = cellData.toString();
						ObjectSetValCh = getCellArray[1];
					} catch (Exception e) {
						update_Report("getcelldata");
					}
					break;

				default:
					actval = "Invalid syntax";
					update_Report("property");
					break;
				}

				if ((ActionVal).equalsIgnoreCase("check")) {
					expval = ObjectSetValCh;
					// Updated on 14.9.2015
					// Specifications change for STH 
					if (objectType.equalsIgnoreCase("radiobutton")) {
						if (expval.equalsIgnoreCase("On")) {
							expval = "True";
						} else if (expval.equalsIgnoreCase("Off")) {
							expval = "False";
						}
					}
					
					if (ObjectSetCh.equalsIgnoreCase("checked") 
							|| ObjectSetCh.equalsIgnoreCase("visible")
							|| ObjectSetCh.equalsIgnoreCase("enabled")
							|| ObjectSetCh.equalsIgnoreCase("exist")) {
						if (expval.equalsIgnoreCase(actval)) {
							// Updated on 14.9.2015
							// Specifications change for STH 
							
							System.out
								.println("Actual value matches with expected value. "
										+ " Actual Value is " + actval 
										+ " and expected value is " + expval);
							
							if (ObjectSetCh.equalsIgnoreCase("getcelldata")) {
	
								update_Report("executed");
							} else {
								update_Report("executed", ObjectSetCh,
										ObjectSetValCh);
							}
							if (capturecheckvalue == true) {
								screenShot(loopnum, TScrowcount, TScname);
							}	
						} else {
							
							System.out
									.println("Actual value doesn't match with expected value. Actual value is "
											+ actval
											+ " expected value is "
											+ expval);
	
							if (ORvalname == "exit") {
								update_Report("missing", ObjectSetCh,
										ObjectSetValCh);
	
							} else {
								if (ObjectSetCh.equalsIgnoreCase("getcelldata")) {
									update_Report("failed");
								} else {
									update_Report("failed", ObjectSetCh,
											ObjectSetValCh);
								}
	
							}
							if (capturecheckvalue == true) {
								screenShot(loopnum, TScrowcount, TScname);
							}
						}
					} else {
						if (expval.equals(actval)) {
							// Updated on 14.9.2015
							// Specifications change for STH 
							System.out
								.println("Actual value matches with expected value. "
									+ " Actual Value is " + actval 
									+ " and expected value is " + expval);
							if (ObjectSetCh.equalsIgnoreCase("getcelldata")) {
	
								update_Report("executed");
							} else {
								update_Report("executed", ObjectSetCh,
										ObjectSetValCh);
							}
							if (capturecheckvalue == true) {
								screenShot(loopnum, TScrowcount, TScname);
							}	
						} else {
							System.out
									.println("Actual value doesn't match with expected value. Actual value is "
											+ actval
											+ " expected value is "
											+ expval);
	
							if (ORvalname == "exit") {
								update_Report("missing", ObjectSetCh,
										ObjectSetValCh);
	
							} else {
								if (ObjectSetCh.equalsIgnoreCase("getcelldata")) {
									update_Report("failed");
								} else {
									update_Report("failed", ObjectSetCh,
											ObjectSetValCh);
								}
							}
							if (capturecheckvalue == true) {
								screenShot(loopnum, TScrowcount, TScname);
							}
						}						
					}
				} else if ((ActionVal).equalsIgnoreCase("storevalue"))

				{
					varname = ObjectSetValCh;

					if (actval.equalsIgnoreCase("Invalid syntax")) {
						update_Report("missing", ObjectSetCh, ObjectSetValCh);

					} else {
						if (map.containsKey(varname)) {
							map.remove(varname);
							map.put(varname, actval);

							if (ObjectSetCh.equalsIgnoreCase("getcelldata")) {
								update_Report("executed");
							} else {
								update_Report("executed", ObjectSetCh,
										ObjectSetValCh);
							}
							System.out
									.println("Overwriting the value of the variable "
											+ varname
											+ " to store the value as mentioned in the test case row number"
											+ rowcnt);
						} else {
							map.put(varname, actval);
							if (ObjectSetCh.equalsIgnoreCase("getcelldata")) {
								update_Report("executed");
							} else {
								update_Report("executed", ObjectSetCh,
										ObjectSetValCh);
							}
							System.out
									.println("Overwriting the value of the variable "
											+ varname
											+ " to store the value as mentioned in the test case row number"
											+ rowcnt);
							if (ObjectSetValCh == "nodatarow") {
								update_Report("missing");
							} else {

							}
						}
					}
					if (capturestorevalue == true) {
						screenShot(loopnum, TScrowcount, TScname);
					}
				}

			} else {
				update_Report("objNotFound");
			}

		} catch (Exception e) {
	

				update_Report("failed", e);
		
		}
	
	}
	
	@After
	public void close() throws Exception

	{
		try {
			System.out.println("Test end.");

			D8.quit();
		} catch (UnhandledAlertException e) {
		
			LOGGER.log(Level.SEVERE, "Because of specification of SeleniumWebDriver, downloading may be failed  "+ e.getMessage() );
			System.out.println(e);
			System.out
					.println("Because of specification of SeleniumWebDriver, downloading may be failed.");
			System.out
					.println("Please confirm the report file and screenshot about test result.");
			LOGGER.log(Level.SEVERE, "Please confirm the report file and screenshot about test result." );
		}
	}
	
	private static void func_FindObj(String strObjtype, String strObjpath)
			throws Exception

	{
		if (strObjpath.contains("@@@")) {
			
			
		
			String eCellDataVal  = TScsheet.getRow(j).getCell(4).getStringCellValue();
			
			
			if (map.containsKey(eCellDataVal)) {
				eCellDataVal = map.get(eCellDataVal);
			}
			strObjpath = strObjpath.replace("@@@", eCellDataVal);

		}

		try {
			if (strObjtype.length() > 0 && strObjpath.length() > 0) {

				if (strObjtype.equalsIgnoreCase("id")) {
					elem = D8.findElement(By.id(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("name")) {
					elem = D8.findElement(By.name(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("xpath")) {
					elem = D8.findElement(By.xpath(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("link")) {
					elem = D8.findElement(By.linkText(strObjpath.toString()));
				} else if (strObjtype.equalsIgnoreCase("css")) {
					elem = D8.findElement(By.cssSelector(strObjpath));
				}

			}
		} catch (Exception ex2) {
			update_Report("objectmiss");
			LOGGER.log(Level.SEVERE, "Error Occured in func_FindObj: " +ex2.getMessage());
			System.out.println(ex2.toString());
			elem = null;
		}
		}
	
	private static WebElement  func_FindObj_return(String strObjtype, String strObjpath)
	 
	{
			 	
		if (strObjpath.contains("@@@")) {
			
		
			String eCellDataVal  = TScsheet.getRow(j).getCell(4).getStringCellValue();
			if (map.containsKey(eCellDataVal)) {
				eCellDataVal = map.get(eCellDataVal);
			}
			strObjpath = strObjpath.replace("@@@", eCellDataVal);

		}

		try {
			if (strObjtype.length() > 0 && strObjpath.length() > 0) {

				
				if (strObjtype.equalsIgnoreCase("id")) {
					elem = D8.findElement(By.id(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("name")) {
					elem = D8.findElement(By.name(strObjpath));
				} else if (strObjtype.equalsIgnoreCase("xpath")) {
				 
					elem = D8.findElement(By.xpath(strObjpath));
					
				} else if (strObjtype.equalsIgnoreCase("link")) {
					elem = D8.findElement(By.linkText(strObjpath.toString()));
				} else if (strObjtype.equalsIgnoreCase("css")) {
					elem = D8.findElement(By.cssSelector(strObjpath));
				}

			}
		} catch (Exception ex2) {
	 
				try {
					update_Report("ObjectNotFound");
				} catch (IOException e) {
					 System.out.println("  Object Not Found Error occured  "+e.getMessage());
					 LOGGER.log(Level.SEVERE, "Object Not Found Error occured in func_FindObj: " +e.getMessage());
					 
				}
			 
			 
			
			elem = null;
		}
return elem;
			
	}
	
	private static boolean search_Excel(String toSearch){

		  boolean isfound=false;
		String[] dCellDataValCh = dCellData.split(":");
 	String ObjectSetCh = dCellDataValCh[0];
 	String ObjectSetValCh = null;
 	boolean columnFound = false;
 	String compare = null;
 	int columnno =0;
	String getColumn = null;
 	if (dCellDataValCh.length == 2) {
			ObjectSetValCh = dCellDataValCh[1];
		}
 	if (ObjectSetValCh.contains("dt_")) {
			iflag = 0;
		
			String ObjectSetValtableheader[] = ObjectSetValCh
					.split("_",2);
			int column = 0;
                         
                          String Searchincolumn = ObjectSetValtableheader[1];
                                             
                         
                     	Row row = ws.getRow(0);
          			
                      int cell = row.getLastCellNum();
                      for(int i=0;i<cell;i++){
                      	getColumn = (ws.getRow(0).getCell(i).getStringCellValue());
                    	  if(getColumn.equalsIgnoreCase(Searchincolumn)){
                    		  columnFound = true;
                    		  break;
                    	  }
                      }                                                            
                                                  
                         if((columnFound == true)&&(columnno!=0)){
                      	   int rowcnt = 0;
                             for(Row r : ws){
                             	rowcnt= ws.getLastRowNum();
                             	
                             }
                        
                             System.out.println("Row Count is  "+rowcnt);
                             
                             for(int i=0;i<rowcnt;i++){
                             	String getCelldata = ws.getRow(i).getCell(columnno).getStringCellValue();
                             	if(getCelldata.equalsIgnoreCase(toSearch)){
                             		System.out.println("Script Found  "+getCelldata);
                             		isfound = true;
                             	}
                             }
                         	 
                    
                         }
                        
		
		
 	}
 	return isfound;
	
		
	}
		
	public static int ifContidionSkipper(String strifConditionStatus)
			throws Exception {
		try {
			String strKeyword;
			int intIfEndConditionCount, intIfConditionCount;
		
			intIfConditionCount = 1;
			intIfEndConditionCount = 0;

			if (strifConditionStatus.equalsIgnoreCase("false")) {
			

				do {
					j = j + 1;
					strKeyword =TScsheet.getRow(j).getCell(1).getStringCellValue();
			 
					 

					if (strKeyword.equalsIgnoreCase("Condition")) {
						intIfConditionCount = intIfConditionCount + 1;
					}

					if (strKeyword.equalsIgnoreCase("Endcondition")) {
						intIfEndConditionCount = intIfEndConditionCount + 1;

						if (intIfConditionCount == intIfEndConditionCount) {
							j = j + 1;
							break;
						}
					}

				} while (true);
			}
		} catch (Exception e) {
			System.out.println(e);
			 LOGGER.log(Level.SEVERE, "Error Occured  in  ifContidionSkipper: " +e.getMessage());
			
		}
		return j;
	}
	
	public String func_IfCondition(String strConditionArgs) throws Exception {

		int iFlag = 1;
		String str[] = strConditionArgs.split(";");
		String value1 = str[0];
		String value2 = str[2];
		String strOperation = str[1];
		strOperation = strOperation.toLowerCase().trim();
		switch (strOperation.toLowerCase()) {
		case "equals":
			if (value1.substring(0, 1).equalsIgnoreCase("#")) {
				 
				value1 = map.get(value1.substring(1, (value1.length())));
				System.out
						.println("Variable used in condition statement has value: "
								+ value1);
				
				
				if (value1.trim().equalsIgnoreCase(value2.trim())) {

					iFlag = 0;
				}
			} else if (value1.trim().equalsIgnoreCase(value2.trim())) {
				iFlag = 0;
			}
			break;

		case "notequals":
			if (value1.substring(0, 1).equalsIgnoreCase("#")) {
				value1 = map.get(value1.substring(1, (value1.length())));
				System.out
						.println("Variable used in condition statement has values: "
								+ value1);
				if (!value1.trim().equalsIgnoreCase(value2.trim())) {

					iFlag = 0;
				}

			} else if (!value1.trim().equals(value2.trim())) {
				iFlag = 0;
			}
			break;
		case "greaterthan":
			if (value1.substring(0, 1).equalsIgnoreCase("#")) {
				value1 = map.get(value1.substring(1, (value1.length())));
				if (isInteger(value1) && isInteger(value2)) {
					if (Integer.parseInt(value1) > Integer.parseInt(value2)) {
						iFlag = 0;
					}
				}
			}

			else if (isInteger(value1) && isInteger(value2)) {
				if (Integer.parseInt(value1) > Integer.parseInt(value2)) {
					iFlag = 0;
				}
			}

			else {
				System.out.println("Give Only Integers for Compare ");
			}
			break;
		case "lessthan":
			if (value1.substring(0, 1).equalsIgnoreCase("#")) {
				value1 = map.get(value1.substring(1, (value1.length())));
				if (isInteger(value1) && isInteger(value2)) {
					if (Integer.parseInt(value1) < Integer.parseInt(value2)) {
						iFlag = 0;
					}
				}
			}

			else if (isInteger(value1) && isInteger(value2)) {
				if (Integer.parseInt(value1) < Integer.parseInt(value2)) {
					iFlag = 0;
				}
			}

			else {
				LOGGER.log(Level.SEVERE, "Error Occured in lessthan : Give Only Integers for Compare" );
				System.out.println("Give Only Integers for Compare ");
			}
			break;
		default:
			update_Report("missing");
			break;
		}
		if (iFlag == 0) {
			 
			return "true";

		} else {
			return "false";
		}
	
	}
	
	public void arrayreSize() {
		if (loopstart.length <= loopsize) {
			loopstart = Arrays.copyOf(loopstart, loopstart.length + 1);
			loopend = Arrays.copyOf(loopend, loopend.length + 1);
			loopcnt = Arrays.copyOf(loopcnt, loopcnt.length + 1);
			dtrownumloop = Arrays.copyOf(dtrownumloop, dtrownumloop.length + 1);
			loopcount = Arrays.copyOf(loopcount, loopcount.length + 1);
		}
	}
	
	public void bCellAction(String scriptName) throws Exception {
		try {
		 
			switch (ActionVal.toLowerCase()) {
			case "loop":
				startrow = j;
				dtrownum = 1;
				loopsize = loopsize + 1;
				if (loopsize >= 1) {
					arrayreSize();

				}
				loopflag = 1;
				loopcount[loopsize] = Integer.parseInt(cCellData);
				loopstart[loopsize] = j;
				loopcnt[loopsize] = 0;
				dtrownumloop[loopsize] = dtrownum;
				update_Report("loop : " + "Start of loop : " + (loopsize + 1));
				update_Report("executed");
				break;
			case "endloop":
				loopcnt[loopsize] = loopcnt[loopsize] + 1;
				loopnum = loopnum + 1;
				if (loopcnt[loopsize] == loopcount[loopsize]) {
					update_Report("loop" + " End of Loop : " + (loopsize + 1)
							+ " : Loop count : " + loopcnt[loopsize]);
					loopsize = loopsize - 1;
					if (loopsize >= 0)
						dtrownum = dtrownumloop[loopsize];
					else {
						dtrownum = 1;
						loopflag = 0;
					}
					update_Report("executed");
				} else {
					j = loopstart[loopsize];
					dtrownum = dtrownum + 1;
					dtrownumloop[loopsize] = dtrownum;
					update_Report("loop" + " End of Loop : " + (loopsize + 1)
							+ " : Loop count : " + loopcnt[loopsize]);
				}
				rowcnt = 1;
				break;
			// 2016-03-23 close Keyword Support
			case "launchapp":
				if(D8 == null) {
					setBrowser(browserName);
				}
				D8.get(cCellData);
				update_Report("executed");
				break;
			// 2016-03-23 close Keyword Support
			case "close":
				D8.close();
				D8 = null;
				update_Report("executed");
				break;

			case "wait":
			 
				Thread.sleep((waitcCellDatanum));
				update_Report("executed");
				break;

			case "condition":
				
				 

				String strConditionStatus = func_IfCondition(cCellData);
				if (strConditionStatus.equalsIgnoreCase("false")) {
					 
					
					conditionline = j;
					j = ifContidionSkipper(strConditionStatus);
					j = j - 1;
				}
				if (strConditionStatus.equalsIgnoreCase("false")) {
				 

					update_Report("condFailed");
				} else {
					

					update_Report("executed");
				}
				break;
				
			case "endcondition":
				
				update_Report("executed");
				break;

			case "screencaptureoption":
				String[] sco = cCellData.split(";");

				for (int s = 0; s < sco.length; s++) {
					if (sco[s].equalsIgnoreCase("perform")) {
						captureperform = true;
					}
					if (sco[s].equalsIgnoreCase("storevalue")) {
						capturestorevalue = true;
					}
					if (sco[s].equalsIgnoreCase("check")) {
						capturecheckvalue = true;
					}

				}
				update_Report("executed");
				break;

			case "importdata":
				try {
					String xcelpath = cCellData;
					FileInputStream fs3 = null;				 
					fs3 = new FileInputStream(new File(xcelpath));			 
					
					XSSFWorkbook DTworkbook = new XSSFWorkbook(fs3);			
					FormulaEvaluator DTevaluator = DTworkbook.getCreationHelper().createFormulaEvaluator();
					
					if(update.equalsIgnoreCase("yes")){
						DTevaluator.evaluateAll();
					}
					DTsheet = DTworkbook.getSheetAt(0);				 
					update_Report("executed");
				} catch (Exception e) {
					LOGGER.log(Level.SEVERE, "Error Occured in importdata : No Data table found  " +e.getMessage() );
					update_Report("nodatatable");
				}
				break;

				
			case "fetchdb":
				
				try{
					
					boolean parameter = false;
					 sqlquery = cCellData;
					 
					 if((database.equalsIgnoreCase("mssql"))||(database.equalsIgnoreCase("mysql"))){
						 isinvaliddb = true;
					 }else{
						 isinvaliddb=false;
						 update_Report("invalidConnection");
						 return;
					 }
				
					 if(host_name.length()==0){
						 System.out.println("Please Enter Host Name");
						 isinvaliddb=false;
						 update_Report("invalidConnection");
						 return;
					 }
					 					 
					 if((schemaname.length()==0)){
						 System.out.println("Please Enter Schema Name");
						 isinvaliddb=false;
						 update_Report("invalidConnection");
						 return;
					 }
					 
					 String substr_query = (sqlquery.trim()).substring(0,6);
					 System.out.println("Query is  "+sqlquery);
			
					 if(sqlquery.length()>0){
							String query = sqlquery.trim();
							String isSelect = query.substring(0, 6);
							 		
						 
									if((! isSelect.equalsIgnoreCase("select"))){
										
									 parameter=true;
									
									}
															
							}else{
								System.out.println("SQL query is not given, Please enter SQL query ");
								update_Report("invalidQuery");
								return;
							}
					 if(parameter == true){
							update_Report("invalidQuery");
							return;
						}
										 
				String url =createConnection(database, schemaname, username, password);
				 if(url.length()>0){
					 isconnected = getConnection(url);		
				 }
				 if(isconnected){
				 		 
				executeQuery(rs);
				update_Report("executed");
				 
				 }else{
					 update_Report("invalidConnection");
					 
					 isinvaliddb = false;
				 }
					
				}catch (Exception e){
					
					System.out.println(e.getMessage());
				 
				}
			 
				break;

				
			case "comparedbcell":
				try{
							 
			if(isinvaliddb){
				ScreenshotTypeFlag = 0;
				WebElement getElement = getWebElement();
				if(getElement==null){
					System.out.println(" No matching Element found in Object Repository");
					update_Report("ObjectNotFound");
					return;
					
				}
				String toSearch = getElement.getText();
			
				boolean isfound = search_Excel(toSearch);
				if(isfound){	
					update_Report("executed");
				}else{
					update_Report("NoMatchinDataTable");
				}
				}else{
					System.out.println("Basic parameters are invalid");
					 update_Report("invalidConnection");
					 
				}
				}
			catch(Exception e){
				LOGGER.log(Level.SEVERE, "Exception Occured in comparedbcell : No Data table found  " +e.getMessage() );
					System.out.println("Exception Occured as -"+e.getMessage());
					
				}
				break;
				
			case "screencapture":
				ScreenshotTypeFlag = 0;
				screenShot(loopnum, TScrowcount, TScname);
				update_Report("executed");
				break;
			case "context":// Updated on 16.12.2013 for page support
				cCellObjName = "";// Updated on 16.12.2013 for page support
				cCellObjType = "";// Updated on 16.12.2013 for page support
				dCellDataVal = null;
				int frameindex = 0;
				if (cCellData.contains(";")) // Updated on 16.12.2013 for page
												// and Dialog Support
				{

					if (cCellData.endsWith(";")) {
						cCellObjType = cCellData.substring(0,
								cCellData.length() - 1);
						cCellObjName = cCellData.substring(0,
								cCellData.length() - 1);

					} else {
						cCellDataVal = cCellData.split(";");
						cCellObjType = cCellDataVal[0];
						cCellObjName = cCellDataVal[1];
					}
				}

				else {
					cCellObjType = cCellData;
					if (cCellObjType.equalsIgnoreCase("frame")
							|| cCellObjType.equalsIgnoreCase("iframe")) {
						frameindex = 1;
					} else {
						cCellObjName = cCellData;
					}
				}
				if (cCellObjType.equalsIgnoreCase("frame")
						|| cCellObjType.equalsIgnoreCase("iframe")) {
					if (frameindex == 1) {
						D8.switchTo().parentFrame();
						update_Report("executed");
						frameindex = 0;
					// 2015-06-17 add <---
					} else if("default".equals(cCellObjName)) {
						D8.switchTo().defaultContent();	
						update_Report("executed");
					} else if(cCellObjName.matches("^[0-9]+") == true){
						int test = new Integer(cCellObjName);
						D8.switchTo().frame(test);
						update_Report("executed");
					// 2015-06-17 add --->
					} else {
						for (int k = 0; k < ORrowcount; k++) {
							if((ORsheet.getRow(k).getCell(1).getStringCellValue()).equalsIgnoreCase(cCellObjName)==true)
						 
							{
								String[] ORcellData = new String[3];
							
								ORcellData = ORsheet.getRow(k).getCell(4).getStringCellValue().split("=");
								
								 
								
								ORvalname = ORcellData[0];
								
								ORvalue = ORsheet.getRow(k).getCell(4).getStringCellValue().substring(ORcellData[0].length()+1);
								
						 
								
								k = ORrowcount + 1;
							}

						}
						func_FindObj(ORvalname, ORvalue);
						if (elem == null) {
							return;
						} else {
							D8.switchTo().frame(elem);
							update_Report("executed");
						}
					}
				} else {

					String windowType = null;
					try {
						dCellData.toString();
						if (dCellData.contains("::")) {
							String tempDCellData = dCellData;
							dCellDataVal = dCellData.split(";");
							windowType = dCellDataVal[0];
							ObjectSetVal = dCellDataVal[1];
							dCellData = dCellData.substring(
									dCellData.indexOf("::") + 2,
									dCellData.length());
							if (dCellData.contains(";")) // Updated on
															// 16.12.2013 for
															// page
							// and Dialog Support
							{
								if (dCellData.endsWith(";")) {
									cCellObjType = dCellData.substring(0,
											dCellData.length() - 1);
									cCellObjName = dCellData.substring(0,
											dCellData.length() - 1);

								} else {
									dCellDataVal = dCellData.split(";");
									cCellObjType = dCellDataVal[0];
									cCellObjName = dCellDataVal[1];

								}
							}

							else {
								cCellObjType = dCellData;

								if (cCellObjType.equalsIgnoreCase("frame")
										|| cCellObjType
												.equalsIgnoreCase("iframe")) {
									frameindex = 1;
								} else {
									cCellObjName = dCellData;
								}
							}
							if (cCellObjType.equalsIgnoreCase("frame")
									|| cCellObjType.equalsIgnoreCase("iframe")) {
								if (frameindex == 1) {
									D8.switchTo().parentFrame();
									dCellData = tempDCellData;
									update_Report("executed");
									frameindex = 0;

								} else {
									for (int k = 0; k < ORrowcount; k++) {
										
										if(((ORsheet.getRow(k).getCell(1).getStringCellValue())						
										.equalsIgnoreCase(cCellObjName) == true)) {
											String[] ORcellData = new String[3];
											
											ORcellData =((ORsheet.getRow(k).getCell(4).getStringCellValue()).split("="));
											
										 
											
											ORvalname = ORcellData[0];
											
										ORvalue = ORsheet.getRow(k).getCell(4).getStringCellValue().substring(ORcellData[0].length()+1);
																				
											
											k = ORrowcount + 1;

										}

									}
									func_FindObj(ORvalname, ORvalue);
									if (elem == null) {
										return;
									} else {
										D8.switchTo().frame(elem);
										dCellData = tempDCellData;

										update_Report("executed");

									}
								}
							}
						} else if (dCellData.contains(";")) {
							if (dCellData.endsWith(";")) {

								windowType = dCellData.toString();
								ObjectSetVal = dCellData.toString();

							} else {
								dCellDataVal = dCellData.split(";");
								windowType = dCellDataVal[0];
								ObjectSetVal = dCellDataVal[1]; // 2015-06-17
								if (ObjectSetVal.substring(0, 1)
										.equalsIgnoreCase("#")) {
									if (map.get(ObjectSetVal.substring(1,
											(ObjectSetVal.length()))) != null) {
										ObjectSetVal = map
												.get(ObjectSetVal
														.substring(
																1,
																(ObjectSetVal
																		.length())));

									} else {
										ObjectSetVal = "";
									}
								// data table check Add 2015-08-19 --->
								} else if (ObjectSetVal.contains("dt_")) {
									String ObjectSetValtableheader[] = ObjectSetVal.split("_");
									
									Row r1 = DTsheet.getRow(0);
									int DTcolumncountCh = r1.getLastCellNum();
									
								 
									int column = 0;
									String Searchtext = ObjectSetValtableheader[1];

									for (column = 0; column < DTcolumncountCh; column++) {
									if(Searchtext.equalsIgnoreCase(DTsheet.getRow(0).getCell(column).getStringCellValue())==true)
										
									
										{
										ObjectSetVal = DTsheet.getRow(dtrownum).getCell(column).getStringCellValue();
										
										 
											
											if (ObjectSetVal.length() == 0) {
												ObjectSetVal = "";
											}
											iflag = 1;
										}
									}	
								}
								// data table check Add 2015-08-19 <---
							}
						} else {
							windowType = dCellData.toString();
							ObjectSetVal = dCellData.toString();
						}
						// 2015-06-23  <--- condition changed�E�Eotext|dialog;�E�E
						if ((!windowType.equalsIgnoreCase("dialog;")) && ((windowType.equalsIgnoreCase("page")
								|| windowType.equalsIgnoreCase("page;")) && !dCellData.contains("::")
								|| !windowType.equalsIgnoreCase("page;WindowRtn;"))) {
						// 2015-06-23 --->
							int windowNums = 0;
							int windowItr = 0;
							WebDriver newWindow = null;
							Set<String> al = new HashSet<String>();
							al = D8.getWindowHandles();
							windowNums = al.size(); // get the number of window
													// handles
							Iterator<String> windowIterator = al.iterator();
							if (windowNums == 1) { 				 
								// Switch the hundle, if number of available hundle is 1.
								String handle = windowIterator.next();
								D8.switchTo().window(handle);
								// Reset Iterator
   							    windowIterator = al.iterator();
							} else {
								// save the current window handle.
								parentWindowHandle = D8.getWindowHandle();
							}
							if (D8.getTitle().toString()

							.equalsIgnoreCase(ObjectSetVal) == true) {
								update_Report("executed"); // 2015-08-19 Add
							} else {
								if (!((ObjectSetVal.equalsIgnoreCase("page") || (ObjectSetVal
										.equalsIgnoreCase("page;"))) || (ObjectSetVal
										.toString() == ""))) {
									if (D8.getTitle().toString()
											.equalsIgnoreCase(ObjectSetVal) == false) {
										for (windowItr = 0; windowItr < windowNums; windowItr++) {

											String windowHandle = windowIterator
													.next();
											newWindow = D8.switchTo().window(
													windowHandle);

											if (newWindow
													.getTitle()
													.toString()
													.equalsIgnoreCase(
															ObjectSetVal)) {

												windowFound = 1;
												update_Report("executed",
														windowType,
														ObjectSetVal);
												break;

											}

										}
										if (windowFound != 1) {

											update_Report("NoWindowFound",
													windowType, ObjectSetVal);
										}

									}
								} else {
									if (windowNums == 1) {
										update_Report("executed");
									} else {
										int winItr1 = 0;
										String windowHandle = null;

										while (winItr1 != windowNums) {
											windowHandle = windowIterator
													.next();
											System.out.println(windowHandle);
											newWindow = D8.switchTo().window(
													windowHandle);
											if (parentWindowHandle
													.equalsIgnoreCase(windowHandle)) {

												if (winItr1 != windowNums - 1) {
													windowHandle = windowIterator
															.next();
													D8.switchTo().window(
															windowHandle);
													update_Report("executed");
													windowFound = 1;
													break;
												} else

												{
													Iterator<String> windowIterator1 = al
															.iterator();
													windowHandle = windowIterator1
															.next();
													D8.switchTo().window(
															windowHandle);
													update_Report("executed");
													windowFound = 1;
													break;

												}
											}

											winItr1++;
										}
										if (windowFound != 1) {
											update_Report("NoWindowFound");
										}

									}
								}

							}
						// 2015-06-18 <---
						} else if(windowType.equalsIgnoreCase("page;WindowRtn;")) {
							D8.switchTo().window(parentWindowHandle);
							update_Report("executed");
						// 2015-06-18 --->
						}

						if ((windowType.equalsIgnoreCase("dialog") || windowType
								.equalsIgnoreCase("dialog;"))) {
							D8.switchTo().alert();
							update_Report("executed");

						}

					} catch (Exception e) {
					 
						update_Report("failed", e);

					}

				}
				break;
			case "check":
				ScreenshotTypeFlag = 0;
				func_StoreCheck();
				break;
			case "storevalue":
				ScreenshotTypeFlag = 0;
				func_StoreCheck();
				break;
			case "upload":

				if (cCellData.toString() == "") {

					update_Report("filePathNotFound2");
					doUploadDownload("abortupload");
				} else {

					if (new File(cCellData.toString()).exists()) {

					 
						doUploadDownload("upload");
					} else {

						update_Report("filePathNotFound");
						doUploadDownload("abortupload");
					}
				}
				break;
			case "download":

				cCellObjName = "";// Updated on 16.12.2013 for page support
				cCellObjType = "";// Updated on 16.12.2013 for page support
				dCellDataVal = null;

				if (cCellData.contains(";")) // Updated on 16.12.2013 for page
												// and Dialog Support
				{
					if (cCellData.endsWith(";")) {
						cCellObjType = cCellData;
						cCellObjName = cCellData;
					} else {
						cCellDataVal = cCellData.split(";");
						cCellObjType = cCellDataVal[0];
						cCellObjName = cCellDataVal[1];

					}
				}

				else {
					cCellObjType = cCellData;
					cCellObjName = cCellData;
				}
				ObjectSet = dCellData.toString();

			
				
				if (ObjectSetVal.contains("dt_"))
					
					DTcolumncount = DTsheet.getRow(0).getLastCellNum();
					 
				if (!(cCellObjType.equalsIgnoreCase("page")
						|| cCellObjType.equalsIgnoreCase("page;")
						|| cCellObjType.equalsIgnoreCase("dialog")
						|| cCellObjType.equalsIgnoreCase("dialog;")
						|| cCellObjType.equalsIgnoreCase("DownloadDialog") || cCellObjType
							.equalsIgnoreCase("uploadDialog"))) // //Updated on
																// 16.12.2013
																// for
																// supporting
																// page and
																// dialog
				{
					for (int k = 0; k < ORrowcount; k++) {
					 

					if(((ORsheet.getRow(k).getCell(1).getStringCellValue()).equalsIgnoreCase(cCellObjName)==true))
				 
						
						
						{
							String[] ORcellData = new String[3];
							
							ORcellData = ORsheet.getRow(k).getCell(4).getStringCellValue().split("=");
							
							                                                                                            
							ORvalname = ORcellData[0];
							ORvalue = ORsheet.getRow(k).getCell(4).getStringCellValue().substring(ORcellData[0].length() + 1);
							
					                                                                                            	 
												
							k = ORrowcount + 1;
						}
					}

				}
				try {
			
					readAttributeforPerform();
					func_FindObj(ORvalname, ORvalue);
					if (elem == null) {
						return;
					} else {
						if (ObjectSet == "") {

							update_Report("filePathNotFound2");
						} else {
							try {

							

								ObjectSet1 = ObjectSet.split("\\\\");
							} catch (Exception e2) {

								System.out.println(e2);
							}
							for (int i = 0; i < ObjectSet1.length - 1; i++) {
								ObjectSet2 = ObjectSet2 + ObjectSet1[i] + "\\";

							}
							if (new File(ObjectSet2.toString()).exists()) {

								doUploadDownload("download");

							} else {

								update_Report("filePathNotFound1");

							}
							ObjectSet2 = "";
							ObjectSet1 = null;
						}

					}

				} catch (Exception e) {
					LOGGER.log(Level.SEVERE, "Error Occured in Download :   " +e.getMessage() );
					System.out.println(e);
				}
				break;

			case "perform":
				 
				ScreenshotTypeFlag = 0;

				cCellObjName = "";// Updated on 16.12.2013 for page support
				cCellObjType = "";// Updated on 16.12.2013 for page support
				dCellDataVal = null;

				if (cCellData.contains(";")) // Updated on 16.12.2013 for page
												// and Dialog Support
				{
					if (cCellData.endsWith(";")) {
						cCellObjType = cCellData;
						cCellObjName = cCellData;
					} else {
						cCellDataVal = cCellData.split(";");
						cCellObjType = cCellDataVal[0];
						cCellObjName = cCellDataVal[1];

					}
				}

				else {
					cCellObjType = cCellData;
					cCellObjName = cCellData;
				}

				dCellData.toString();

				if (dCellData.contains(":")) {
					if (!dCellData.endsWith(":")) {
						dCellDataVal = dCellData.split(":");
						ObjectSet = dCellDataVal[0].toLowerCase();
						ObjectSetVal = dCellDataVal[1];

						if (dCellDataVal.length > 2) {
							for (int dCellvalcount = 2; dCellvalcount < dCellDataVal.length; dCellvalcount++) {
								if (ObjectSetVal.startsWith("dt_")
										|| ObjectSetVal.startsWith("#"))

								{
									update_Report("tooManyArguments");
									break;
								} else {
									ObjectSetVal = ObjectSetVal + ":"
											+ dCellDataVal[dCellvalcount];
								}
							}
						}
					} else {
						ObjectSet = dCellData.replace(":", "");
						ObjectSetVal = "";

					}
				} else {
					ObjectSet = dCellData.toString();
				}
				DTcolumncount = 0;
				if (ObjectSetVal.contains("dt_"))
					DTcolumncount = DTsheet.getRow(0).getLastCellNum();
					//DTcolumncount = DTsheet.getColumns();
				if (!(cCellObjType.equalsIgnoreCase("page")
						|| cCellObjType.equalsIgnoreCase("page;")
						|| cCellObjType.equalsIgnoreCase("dialog")
						|| cCellObjType.equalsIgnoreCase("dialog;")
						|| cCellObjType.equalsIgnoreCase("DownloadDialog") || cCellObjType
							.equalsIgnoreCase("uploadDialog"))) // //Updated on
																// 16.12.2013
																// for
																// supporting
																// page and
																// dialog
				{
					for (int k = 0; k < ORrowcount; k++) {
					 
                 
						if(((ORsheet.getRow(k).getCell(1).getStringCellValue()).equalsIgnoreCase(cCellObjName)==true))
						
					                                                                                                  	 
						
						{
							String[] ORcellData = new String[3];
							
							ORcellData = ORsheet.getRow(k).getCell(4).getStringCellValue().split("=",2);
							
						                                                                                                              

							ORvalname = ORcellData[0];
							
							ORvalue =ORsheet.getRow(k).getCell(4).getStringCellValue().substring(ORcellData[0].length() + 1);
							
							                                                                                                     
							k = ORrowcount + 1;
						}
					}

				}

				dCellAction();
				break;
			case "callfunction":

				try {
					func_InvokeFunction(cCellData, dCellData);
				} catch (Exception e) {
					update_Report("userdefined");
				}
				break;
			case "callaction":
				
				if(cCellData.toString().equalsIgnoreCase(ScriptName)){
					System.out.println(" Error Occured. Currently Executing file name should not be provide in CallAction Keyword.");
					System.out.println("Now Executing  Test Script file name "+ScriptName+" which is Equals to script file name "+cCellData+" set in callaction keyword");
					update_Report("fail");
					update_Report("callactionff");
					break;
				}else if(!(cCellData.toString().equalsIgnoreCase(ScriptName))){

				reporttype = 1;
				exeStatus = "Pass";
				String ComponentPath = reusableComponents + cCellData;
				if (cCellData.contains("xlsx")) {
					String ComponentName = cCellData.split(".xlsx")[0];
				//	System.out.println("ComponentName is "+ComponentName);
					XSSFSheet  TestScriptSheet = TScsheet;
					
					FileInputStream ComponentFile1 = null;
			 

					try {
						ComponentFile1 = new FileInputStream(new File(
								ComponentPath));
					                                                                                                               
						XSSFWorkbook  ComponentWorkBook = new XSSFWorkbook(ComponentFile1);
				
						                                                                                             
						
						XSSFSheet ComponentSheet = ComponentWorkBook.getSheetAt(0);
						int ComponentRowCount = 0;
						int introwcnt = 0;
						int introwcntStore = j;
						update_Report("executed");
						j = j + 1;
						TScsheet = ComponentSheet;
						Stack<String> ComponentStack = new Stack<String>();
						globalCompName = ComponentName;
						ComponentStack.push(ComponentName);
						update_Report("callactionstart");
						
						
						ComponentRowCount = ComponentSheet.getLastRowNum();

                                                                                                          						            
						introwcnt = 0;
						for (int jloop = 0; jloop < ComponentRowCount; jloop++) {
							introwcnt = introwcnt + 1;
							j = jloop;
							String CTValidate = "r";
							
							if(((ComponentSheet.getRow(jloop).getCell(0).getStringCellValue()).equalsIgnoreCase(CTValidate)==true))
						                                                                                                                  
							{
								
								// 2016-03-23 Get Cell Value Support(xlsx)
								Action = getCellData(ComponentWorkBook, ComponentSheet.getRow(jloop).getCell(1));
								cCellData = getCellData(ComponentWorkBook, ComponentSheet.getRow(jloop).getCell(2));
								dCellData=	getCellData(ComponentWorkBook, ComponentSheet.getRow(jloop).getCell(3));
                                                                                                    
								String ORPath = objectRepository;
								FileInputStream ComponentFile2 = null;
							 
								try {
									ComponentFile2 = new FileInputStream(
											new File(ORPath));
									 
								} catch (Exception e) {
									System.out.println(" The Action(TestScript) name given is not available in the given path. Check the file path and action name");
								}
								try {
								                                                                                                 
									XSSFWorkbook ORworkbook = new XSSFWorkbook(ComponentFile2);
									ORsheet = ORworkbook.getSheetAt(0);
									ORrowcount = ORsheet.getLastRowNum();
									ActionVal = Action.toLowerCase();
									iflag = 0;
								} catch (Exception e) {

									fail("Excel file of Open2Test is not correct.");
								}
								bCellAction(scriptName);
							/*	System.out.println(Action + "||" + cCellData
										+ "||" + dCellData);*/
								System.out.println((j+1)+" | "+Action+" | "+cCellData+" | "+dCellData+" | "+comments);
							

								jloop = j;
								 
							
							}// End of Execution
						}// End of If that get all rows in Test Script
						System.out.println("The Action(TestScript) name given is not available in the given path. Check the file path and action name");
						update_Report("callactionend");
						globalCompName = ComponentStack.pop();
						
						j = introwcntStore;
						reporttype = 0;
						TScsheet = TestScriptSheet;
					
					} catch (FileNotFoundException FNF1) {
						System.out.println("The Action(TestScript) name given is not available in the given path. Check the file path and action name");
						update_Report("callactionfnf");
					}
				} 
				}
				else {
					System.out.println("The Action(TestScript) name given is not available in the given path. Check the file path and action name");
					update_Report("fail");
					update_Report("callactionff");
					
				}
				break;
			default:
				update_Report("keyword");
				break;

			}
		}

		catch (Exception ex) {
			update_Report("failed", ex);
			LOGGER.log(Level.SEVERE, "failed  " +ex.getMessage() );
			 
			LOGGER.log(Level.SEVERE, "------Error Information : Open2Test------- "  );
			LOGGER.log(Level.SEVERE, "Current Script:" + scriptName );
			LOGGER.log(Level.SEVERE, "Current ScriptPath:" + testScript );
			LOGGER.log(Level.SEVERE, "Using ObjectRepositoryPath:" + objectRepository );
			LOGGER.log(Level.SEVERE, "Current Keyword:" + Action );
			LOGGER.log(Level.SEVERE, "Current ObjectDetails:" + cCellData );
			LOGGER.log(Level.SEVERE, "Current ObjectDetailsPath:" + ORvalue );
			LOGGER.log(Level.SEVERE, "Current Action:" + dCellData );
			LOGGER.log(Level.SEVERE, "------Error Information : Open2Test-------" );
			
			
			System.out.println("------Error Information : Open2Test-------");
			System.out.println("Current Script:" + scriptName);
			System.out.println("Current ScriptPath:" + testScript);
			System.out
					.println("Using ObjectRepositoryPath:" + objectRepository);
			System.out.println("Current Keyword:" + Action);
			System.out.println("Current ObjectDetails:" + cCellData);
			System.out.println("Current ObjectDetailsPath:" + ORvalue);
			System.out.println("Current Action:" + dCellData);
			System.out.println("------Error Information : Open2Test-------");
			
			fail("Cannot test normally by Open2Test.");
			// return;
		}
	}
	
	public static String getCellData(XSSFWorkbook xWB, XSSFCell tCell) {
		String cellStr = null;
		switch (tCell.getCellType()) {
			case XSSFCell.CELL_TYPE_STRING :
				cellStr = tCell.getStringCellValue();
				break;
			case XSSFCell.CELL_TYPE_NUMERIC :
				double dNum = tCell.getNumericCellValue();
				int iNum = (int) dNum;
				cellStr = String.valueOf(iNum);
				break;
			case XSSFCell.CELL_TYPE_FORMULA :
				FormulaEvaluator evaluator = xWB.getCreationHelper().createFormulaEvaluator();
				evaluator.evaluateFormulaCell(tCell);
				cellStr = tCell.getStringCellValue();
				break;
			case XSSFCell.CELL_TYPE_BOOLEAN :
				cellStr = String.valueOf(tCell.getBooleanCellValue());
			default:
				cellStr = "";
		}
		return cellStr;
	}
	
	public static WebElement getWebElement(){
	     	try{
	String[] cCellDataValCh = cCellData.split(";");
	String objectType = cCellDataValCh[0];
	String ObjectValCh = "";
             ObjectValCh = cCellDataValCh[1];
             

         	for (int k = 0; k < ORrowcount; k++) {
         				                                    	 

         		if(((ORsheet.getRow(k).getCell(1).getStringCellValue()).equalsIgnoreCase(ObjectValCh)==true))
         					
         		
         	                                                                                                                      	 
         		{
         						String[] ORcellData = new String[3];
         						
         						ORcellData = (ORsheet.getRow(k).getCell(4).getStringCellValue().split("=",2));
         						
         						                                                                                                   

         						ORvalname = ORcellData[0]; 
         						ORvalue = ORcellData[1]; 
         						k = ORrowcount + 1;
         						objFoundFlag = 1;
         					}

         				}
         
				try {
					WebElement elem;
					elem = func_FindObj_return(ORvalname, ORvalue);
					
					return elem ;
				} catch (Exception e) {
			System.out.println("Error Occured While getting Element-  "+e.getMessage());
			LOGGER.log(Level.SEVERE, "Error Occured While getting Element- " +e.getMessage());
					 
				}
			 
	}catch(Exception e){
		System.out.println("Exception Occured as -"+e.getMessage());
		LOGGER.log(Level.SEVERE, "Exception Occured as- " +e.getMessage());
	}
	return elem ;

}
	 
	public void readAttributeforPerform() throws Exception {


		try {
			if (ObjectSetVal.length() > 0)

			{
				                                                                                                 // //Function was updated on 16.12.2013 for supporting page,
				if (ObjectSetVal.substring(0, 1).equalsIgnoreCase("#")) {

					if (map.get(ObjectSetVal.substring(1,
							(ObjectSetVal.length()))) != null) {
						ObjectSetVal = map.get(ObjectSetVal.substring(1,
								(ObjectSetVal.length())));
					} else {
						ObjectSetVal = "";
					}

				} else if (ObjectSetVal.contains("dt_")) {
					iflag = 0;

					String ObjectSetValtableheader[] = ObjectSetVal.split("_");
					int column = 0;
					String Searchtext = ObjectSetValtableheader[1];

					for (column = 0; column < DTcolumncount; column++) {
						
						if(Searchtext.equalsIgnoreCase(DTsheet.getRow(0).getCell(column).getStringCellValue())==true)
						
						                                                                                                                                     
							{
							ObjectSetVal = DTsheet.getRow(dtrownum).getCell(column).getStringCellValue();
							
							                                                                                                                                    
										
							if (ObjectSetVal.length() == 0) {
								ObjectSetVal = "";
							}
							iflag = 1;

						}

					}

					if (iflag != 1) {
						ObjectSetVal = "nodatarow";
					}

					else {
						update_Report("toomanyarguments");
					}
				}

			}
		} catch (Exception e) {
           System.out.println(" Error Occured while Reading Attribute for Perform  "+e.getMessage());
       	LOGGER.log(Level.SEVERE, "Exception Occured while Reading Attribute for Perform- " +e.getMessage());
		}
	
		
	}
	
	public void dCellAction() throws Exception {

		if (!(cCellObjType.equalsIgnoreCase("page")
				|| cCellObjType.equalsIgnoreCase("page;")
				|| cCellObjType.equalsIgnoreCase("dialog")
				|| cCellObjType.equalsIgnoreCase("dialog;")
				|| cCellObjType.equalsIgnoreCase("downloaddialog")
				|| cCellObjType.equalsIgnoreCase("downloaddialog;")
				|| cCellObjType.equalsIgnoreCase("uploaddialog") || cCellObjType
					.equalsIgnoreCase("uploaddialog;")))// Updated on 16.12.2013
														// for page and dialog
														// support
		{
			try {
				String[] cCellDataValCh = cCellData.split(";");
				String objectType = cCellDataValCh[0];
				readAttributeforPerform();
				func_FindObj(ORvalname, ORvalue);
				if (elem == null) {
					return;
				} else {
					switch (ObjectSet.toLowerCase()) {
					case "set":
						if (objectType.equalsIgnoreCase("textbox")
								|| objectType.equalsIgnoreCase("textarea")) {

							elem.clear();
							StringBuffer inputvalue = new StringBuffer();
							if (ObjectSetVal == "nodatarow") {
								update_Report("missing");
							} else {

								inputvalue.append(ObjectSetVal);
						 

								D8.executeScript(
										"arguments[0].value=arguments[0].value + '"
												+ inputvalue.toString() + "';",
										elem);
							 
								update_Report("executed", ObjectSet,
										ObjectSetVal);
							}

							if (captureperform == true) {
								screenShot(loopnum, TScrowcount, TScname);
							}
						} else {
							update_Report("action1");
						}
						break;
					case "listselect":
						if (objectType.equalsIgnoreCase("listbox")) {
							int foundCount = 0;
							String[] listvalues = ObjectSetVal.split(":");
							List<WebElement> listboxitems = elem
									.findElements(By.tagName("option"));
							Select chooseoptn = new Select(elem);
							chooseoptn.deselectAll();
							if (ObjectSetVal == "nodatarow") {
								update_Report("missing");
							} else {
								for (WebElement opt : listboxitems) {
									for (int i = 0; i < listvalues.length; i++) {
										if (opt.getText().equalsIgnoreCase(
												listvalues[i])) {
											chooseoptn.selectByVisibleText(opt
													.getText());
											foundCount = foundCount + 1;
										}

									}
								}
								if (foundCount == 0
										&& ObjectSetVal.contains(""))

								{
									update_Report("NoBlankAvailable");

								} else {
									update_Report("executed", ObjectSet,
											ObjectSetVal);
								}

								if (captureperform == true) {
									screenShot(loopnum, TScrowcount, TScname);
								}
							}
						} else {
							update_Report("action1");
						}

						break;

					case "select":
						if (objectType.equalsIgnoreCase("combobox")) {
							if (ObjectSetVal != "") {
								new Select(elem)
										.selectByVisibleText(ObjectSetVal);
								update_Report("executed", ObjectSet,
										ObjectSetVal);
							} else if (ObjectSetVal == "nodatarow") {
								update_Report("missing");
							}

							else {
								if (new Select(elem).getOptions().toString()
										.contains("") == true) {
									try {
										new Select(elem)
												.selectByVisibleText("");
										update_Report("executed", ObjectSet,
												ObjectSetVal);
									} catch (Exception e) {
										update_Report("NoBlankAvailable");
									}
								} else {
									update_Report("NoBlankAvailable");
								}
							}
							if (captureperform == true) {
								screenShot(loopnum, TScrowcount, TScname);

							}
						} else {
							update_Report("action1");
						}

						break;

					case "check":

						if (ObjectSetVal == "nodatarow") {
							update_Report("missing");
						} else {

							if (objectType.equalsIgnoreCase("checkbox")) {
								if (elem.isSelected()
										&& ObjectSetVal.equalsIgnoreCase("On")) {
									update_Report("executed", ObjectSet,
											ObjectSetVal);

								} else if ((elem.isSelected() && ObjectSetVal
										.equalsIgnoreCase("off"))) {
									elem.click();
									update_Report("executed", ObjectSet,
											ObjectSetVal);
								} else if (ObjectSetVal.equalsIgnoreCase("on")) {
									elem.click();
									update_Report("executed", ObjectSet,
											ObjectSetVal);
								} else if ((ObjectSetVal
										.equalsIgnoreCase("off"))) {
									update_Report("executed", ObjectSet,
											ObjectSetVal);
								} else {
									update_Report("failed", ObjectSet,
											ObjectSetVal);
								}

								if (captureperform == true) {

									screenShot(loopnum, TScrowcount, TScname);
								}
							} else {
								update_Report("action1");
							}
						}
						break;

					case "click":
						parentWindowHandle1 = D8.getWindowHandle();
						try {

							if (elem.getAttribute("type").toLowerCase()
									.equals("file")
									&& browserName.equalsIgnoreCase("ff")) {

								JavascriptExecutor executor = (JavascriptExecutor) D8;
								executor.executeScript("arguments[0].click();",
										elem);

								update_Report("executed");
								break;

							} else if (elem.getAttribute("type").toLowerCase()
									.equals("file")
									&& browserName.equalsIgnoreCase("ie")
									&& Integer.parseInt(browserver) == 8) {
								update_Report("executed");

							} else {
								elem.click();
								update_Report("executed");
							}
						} catch (Exception exp1) {
							elem.click();
							update_Report("executed");
						}

						if (captureperform == true) {
							screenShot(loopnum, TScrowcount, TScname);
						}
						break;
					case "hover":
						JavascriptExecutor js = (JavascriptExecutor) D8;
						String mouseOverScript = "if(document.createEvent){var evObj = document.createEvent('MouseEvents');evObj.initEvent('mouseover', true, false); arguments[0].dispatchEvent(evObj);} else if(document.createEventObject) { arguments[0].fireEvent('onmouseover');}";
						js.executeScript(mouseOverScript, elem);
						update_Report("executed");
						break;

					case "altclick":
						JavascriptExecutor executor = (JavascriptExecutor) D8;
						executor.executeScript("arguments[0].click();", elem);
						update_Report("executed"	);
						break;
					case "enter":
						elem.sendKeys(org.openqa.selenium.Keys.ENTER);
						update_Report("executed");
						if (captureperform == true) {
							screenShot(loopnum, TScrowcount, TScname);
						}
						break;
					case "setdate":
					 
						Robot robot1 = new Robot();
						ScreenshotTypeFlag = 1;
						String calstring = cCellObjName.toLowerCase();
						if (cCellObjType.equalsIgnoreCase("calendar")
								&& calstring.startsWith("cal_")) {
							try {

								String[] datearray = ObjectSetVal.split("-");
								mm = datearray[0];
								dd = datearray[1];
								yyyy = datearray[2];
								if (Integer.parseInt(mm) > 12
										|| Integer.parseInt(mm) < 1
										|| Integer.parseInt(yyyy) < 1000
										|| Integer.parseInt(yyyy) > 2999
										|| Integer.parseInt(yyyy) < 1700
										|| ((Integer.parseInt(dd) > 30) && (Integer
												.parseInt(dd) == 2
												|| Integer.parseInt(dd) == 4
												|| Integer.parseInt(dd) == 6
												|| Integer.parseInt(dd) == 9 || Integer
												.parseInt(dd) == 11))
										|| (Integer.parseInt(dd) > 28
												&& Integer.parseInt(mm) == 2 && (Integer
												.parseInt(yyyy) % 4 != 0))) {
									update_Report("invaliddate1");
								} else {
									selectDate(mm, dd, yyyy);
								}

							} catch (Exception e) {
								update_Report("invaliddate1");
								robot1.keyPress(KeyEvent.VK_ESCAPE);
								robot1.keyRelease(KeyEvent.VK_ESCAPE);

							}
						} else {

							update_Report("calendaraction");
						}
						break;

					default:
						update_Report("action");
						break;
					}
				}

			} catch (Exception e) {
				System.out.println(e.toString());
				update_Report("failed", e);
			}
		} else {
			try {
				switch (ObjectSet.toLowerCase()) {

				case "ok":// Updated on 16.12.2013 for dialog support
					ScreenshotTypeFlag = 1;
					if (cCellObjType.equalsIgnoreCase("dialog")
							|| cCellObjType.equalsIgnoreCase("dialog;")) {
						dialogSwitch = D8.switchTo().alert();
						dialogSwitch.accept();
						update_Report("executed");
					}
					if (captureperform == true) {
						screenShot(loopnum, TScrowcount, TScname);
					}

					break;
				case "cancel":// Updated on 16.12.2013 for dialog support
					ScreenshotTypeFlag = 1;

					if (cCellObjType.equalsIgnoreCase("dialog")
							|| cCellObjType.equalsIgnoreCase("dialog;")) {
						dialogSwitch = D8.switchTo().alert();
						dialogSwitch.dismiss();
						update_Report("executed");
					}

					if (captureperform == true) {
						screenShot(loopnum, TScrowcount, TScname);
					}
					break;

				case "close":// Updated on 16.12.2013 for dialog and page
								// support
					ScreenshotTypeFlag = 1;

					if (cCellObjType.equalsIgnoreCase("dialog")
							|| cCellObjType.equalsIgnoreCase("dialog;")) {

						dialogSwitch = D8.switchTo().alert();
						dialogSwitch.dismiss();
						update_Report("executed");
						if (captureperform == true) {
							screenShot(loopnum, TScrowcount, TScname);
						}

					}

					else {
						windowFound = 0;
						int windowNums = 0;
						int windowItr = 0;
						String currentWindowHandle = D8.getWindowHandle();
						WebDriver newWindow = null;
						Set<String> al = new HashSet<String>();
						al = D8.getWindowHandles();
						windowNums = al.size();
						Iterator<String> windowIterator = al.iterator();
						if (cCellObjName.equalsIgnoreCase("page;")
								|| cCellObjName.equalsIgnoreCase("page")) {
							if (windowNums == 1) {
								if (captureperform == true) {
									screenShot(loopnum, TScrowcount, TScname);
								}
								D8.close();
								update_Report("executed");
								windowFound = 1;
							} else {
								int winItr1 = 0;
								String windowHandle = null;
								String tempWindowHandle = null;
								while (winItr1 != windowNums) {
									tempWindowHandle = windowHandle;
									windowHandle = windowIterator.next();
									newWindow = D8.switchTo().window(
											windowHandle);
									if (currentWindowHandle
											.equalsIgnoreCase(windowHandle)) {
										if (winItr1 == 0) {
											if (captureperform == true) {
												screenShot(loopnum,
														TScrowcount, TScname);
											}
											D8.close();
											windowHandle = windowIterator
													.next();
											D8.switchTo().window(windowHandle);
											update_Report("executed");
											windowFound = 1;
											break;
										} else {
											if (captureperform == true) {
												screenShot(loopnum,
														TScrowcount, TScname);
											}
											D8.close();
											D8.switchTo().window(
													tempWindowHandle);
											update_Report("executed");
											windowFound = 1;
											break;

										}
									}

									winItr1++;
								}
							}
						} else {
							if (windowNums == 1) {
								if (D8.getTitle().toString()
										.equalsIgnoreCase(ObjectSetVal) == true) {
									if (captureperform == true) {
										screenShot(loopnum, TScrowcount,
												TScname);
									}
									D8.close();
									update_Report("executed");
									windowFound = 1;
								}

							} else {
								if (cCellObjType.equalsIgnoreCase("page")
										&& D8.getTitle().equalsIgnoreCase(
												cCellObjName) == false) {
									for (windowItr = 0; windowItr < windowNums; windowItr++) {
										windowHandle = windowIterator.next();
										newWindow = D8.switchTo().window(
												windowHandle);
										if (newWindow.getTitle()
												.equalsIgnoreCase(cCellObjName)) {
											if (captureperform == true) {
												screenShot(loopnum,
														TScrowcount, TScname);
											}
											newWindow.close();
											update_Report("executed");
											D8.switchTo().window(
													currentWindowHandle);
											windowFound = 1;
											break;
										}

									}

								} else {
									if (cCellObjType.equalsIgnoreCase("page")
											&& D8.getTitle()
													.toString()
													.equalsIgnoreCase(
															cCellObjName) == true) {
								
										int winItr1 = 0;
										String windowHandle = null;
										String tempWindowHandle = null;
										while (winItr1 != windowNums) {
											tempWindowHandle = windowHandle;
											windowHandle = windowIterator
													.next();
											newWindow = D8.switchTo().window(
													windowHandle);
											if (currentWindowHandle
													.equalsIgnoreCase(windowHandle)) {
												if (winItr1 == 0) {
													if (captureperform == true) {
														screenShot(loopnum,
																TScrowcount,
																TScname);
													}
													D8.close();
													windowHandle = windowIterator
															.next();
													D8.switchTo().window(
															windowHandle);
													update_Report("executed");
													windowFound = 1;
													break;
												} else {
													D8.close();
													D8.switchTo().window(
															tempWindowHandle);
													update_Report("executed");
													windowFound = 1;
													break;

												}
											}

											winItr1++;
										}

									}
								}
							}
						}
						if (windowFound != 1) {
							update_Report("NoWindowFound");
						}
					}
					break;
				case "closeupload":
					ScreenshotTypeFlag = 1;
					doUploadDownload("closeupload");
					break;

				case "cancelupload":
					ScreenshotTypeFlag = 1;
					doUploadDownload("cancelupload");
					break;

				default:
					System.out.println("Action not supported");
					break;

				}
			}

			catch (NoAlertPresentException e) {
				LOGGER.log(Level.SEVERE, "Exception Occured in dCellAction- " +e.getMessage());
			}

		}

	}
	
	private void doUploadDownload(String action1) throws Exception {
		// Robot robot = new Robot();
		if (browserName.equalsIgnoreCase("ff")) {
			switch (action1) {
			case "upload":
				try {

					Thread.sleep(2000);
					Runtime.getRuntime().exec(
							execpath + " 2 upload " + cCellData + " "
									+ browserName.toLowerCase());
					update_Report("executed");
				} catch (IOException e) {

					// TODO Auto-generated catch block
					System.out.println(e);

				}
				break;

			case "closeupload":
				try {

					Runtime.getRuntime().exec(
							execpath + " 2 closeupload " + cCellData + " "
									+ browserName.toLowerCase());
					update_Report("executed");

				} catch (IOException e) {

					 
					System.out.println(e);
					LOGGER.log(Level.SEVERE, "Exception Occured in closeupload- " +e.getMessage());

				}
				break;
			case "abortupload":
				try {

					Runtime.getRuntime().exec(
							execpath + " 2 closeupload " + "abort" + " "
									+ browserName.toLowerCase());

				} catch (IOException e) {

					LOGGER.log(Level.SEVERE, "Exception Occured in abortupload- " +e.getMessage());
					System.out.println(e);

				}
				break;

			case "cancelupload":
				try {

					Runtime.getRuntime().exec(
							execpath + " 2 cancelupload " + cCellData + " "
									+ browserName.toLowerCase());
					update_Report("executed");
				} catch (IOException e) {

					LOGGER.log(Level.SEVERE, "Exception Occured in cancelupload- " +e.getMessage());
					System.out.println(e);

				}
				break;

			case "download":
				try {

					Runtime.getRuntime().exec(
							execpath + " 2 download " + ObjectSet + " "
									+ browserName.toLowerCase() + " "
									+ elem.getAttribute("href"));
					Thread.sleep(4000);
					update_Report("executed");
				} catch (IOException e) {
                   System.out.println("Error Occured in Dowload keyword execution  "+e.getMessage());
               	LOGGER.log(Level.SEVERE, "IOException Occured in download- " +e.getMessage()); 

				}
				break;

			default:
				System.out.println("Action not supported");
				break;
			}
		} else if (browserName.equalsIgnoreCase("ie")) {
			switch (action1) {
			case "upload":
				if (Integer.parseInt(browserver) != 8) {
					try {

						Runtime.getRuntime().exec(
								execpath + " 2 upload " + cCellData + " "
										+ browserName.toLowerCase());
						update_Report("executed");
					} catch (IOException e) {

					 	LOGGER.log(Level.SEVERE, "IOException Occured in upload- " +e.getMessage()); 
						System.out.println(e);

					}

				} else {
					elem.sendKeys(cCellData);
				}
				break;

			case "closeupload":
				try {

					Runtime.getRuntime().exec(
							execpath + " 2 closeupload " + cCellData + " "
									+ browserName.toLowerCase());
					update_Report("executed");

				} catch (IOException e) {

				 	LOGGER.log(Level.SEVERE, "IOException Occured in closeupload- " +e.getMessage()); 
					System.out.println(e);

				}
				break;

			case "cancelupload":
				try {

					Runtime.getRuntime().exec(
							execpath + " 2 cancelupload " + cCellData + " "
									+ browserName.toLowerCase());
					update_Report("executed");
				} catch (IOException e) {

					LOGGER.log(Level.SEVERE, "IOException Occured in cancelupload- " +e.getMessage()); 
					System.out.println(e);

				}
				break;

			case "download":

				try {
					Runtime.getRuntime().exec(
							execpath + " 2 download " + ObjectSet + " "
									+ browserName.toLowerCase() + " "
									+ elem.getAttribute("href"));
					update_Report("executed");
				} catch (IOException e) {
					System.out.println("IO exception occured in   download key woard "+e.getMessage());
					LOGGER.log(Level.SEVERE, "IOException Occured in download- " +e.getMessage()); 

				}
				break;

			}

		}

		if (captureperform == true) {
			screenShot(loopnum, TScrowcount, TScname);
		}
		 
	}
	
	public void selectDate(String mm2, String dd2, String yyyy2)
			throws Exception {
		String dateClass = null;
		int objectfound = 0;
		int monthnum1 = 0;
		String usedTitleMonth = null;
		String usedTitleYear = null;
		String usedTitletag = null;
		WebElement prevMonth = null;
		WebElement nextMonth = null;
		WebElement titleYear = null;
		WebElement daytoClick = null;
		int titleYearnum = 0;
		String titleMonthString = "";
		WebElement titleMonth = null;
		String datePickerType = "";
		String locatorNext = "";
		String locatorPrev = "";
		String outerHTMLCalendar = "";
		outerHTMLCalendar = elem.getAttribute("outerHTML");
		elem.click();
		int setmm = Integer.parseInt(mm2);
		int setyyyy = Integer.parseInt(yyyy2);
		int setdd = Integer.parseInt(dd2);
		WebElement titleDefault = null;
		String titleDefaultText;
	 
		List<WebElement> realCategoryDeciders = new ArrayList<WebElement>();
 
		Robot robot1 = new Robot();
		int isDateNotSelected = 0;
		int isDateSelected = 0;

		if (outerHTMLCalendar.toLowerCase().contains("datepicker")) {
			datePickerType = "jquery";
		} else if (outerHTMLCalendar.toLowerCase().contains("type=")
				&& outerHTMLCalendar.toLowerCase().contains("date")) {
			datePickerType = "html5";
		} else {
			datePickerType = "boot";
		}

		switch (setmm) {

		case 1:
			monthpart = "Jan";
			monthpartjap = "1月";
			break;
		case 2:
			monthpart = "Feb";
			monthpartjap = "2月";
			break;
		case 3:
			monthpart = "Mar";
			monthpartjap = "3月";
			break;
		case 4:
			monthpart = "Apr";
			monthpartjap = "4月";
			break;
		case 5:
			monthpart = "May";
			monthpartjap = "5月";
			break;
		case 6:
			monthpart = "Jun";
			monthpartjap = "6月";
			break;
		case 7:
			monthpart = "Jul";
			monthpartjap = "7月";
			break;
		case 8:
			monthpart = "Aug";
			monthpartjap = "8月";
			break;
		case 9:
			monthpart = "Sep";
			monthpartjap = "9月";
			break;
		case 10:
			monthpart = "Oct";
			monthpartjap = "10月";
			break;
		case 11:
			monthpart = "Nov";
			monthpartjap = "11月";
			break;
		case 12:
			monthpart = "Dec";
			monthpartjap = "12月";
			break;
		default:
			update_Report("invalidmonth");
			break;
		}
		int titleLength;
		switch (datePickerType) {

		case "jquery":
			if (!(lastSelecteddateJQ == setdd && lastSelectedmonthJQ == setmm && lastSelectedyearJQ == setyyyy)) {

				for (String strPrevMonth : setObj.envprevMonth1) {
					try {
						prevMonth = D8
								.findElementByXPath("//a[contains(@class,'"
										+ strPrevMonth + "')]");
						locatorPrev = prevMonth.getAttribute("class")
								.toString();
						objectfound = 1;
						break;
					} catch (Exception e) {
						continue;
					}
				}
				if (objectfound == 0) {
					update_Report("prevmonth");
					break;
				} else {
					objectfound = 0;
				}
				for (String strNextMonth : setObj.envnextMonth1) {
					try {
						nextMonth = D8
								.findElementByXPath("//a[contains(@class,'"
										+ strNextMonth + "')]");
						locatorNext = nextMonth.getAttribute("class")
								.toString();
						objectfound = 1;
						break;
					} catch (Exception e) {
						continue;
					}
				}

				if (objectfound == 0) {
					update_Report("nextmonth");

					break;
				} else {
					objectfound = 0;

				}

				for (String strtitleMonth : setObj.envtitleMonth) {

					try {
						titleMonth = D8
								.findElementByXPath("//span[contains(@class,'"
										+ strtitleMonth + "')]");
						titleMonthString = titleMonth.getText();

						usedTitleMonth = strtitleMonth;
						usedTitletag = "span";
						objectfound = 1;
						break;
					} catch (Exception e) {
						try {
							titleMonth = D8
									.findElementByXPath("//select[contains(@class,'"
											+ strtitleMonth + "')]");
							titleMonthString = new Select(titleMonth)
									.getFirstSelectedOption().getText()
									.toString();
							usedTitleMonth = strtitleMonth;
							usedTitletag = "select";
							objectfound = 1;
							break;
						} catch (Exception e4) {

						}
						continue;
					}

				}

				if (objectfound == 0) {
					update_Report("titlemonth");

					break;
				} else {
					objectfound = 0;
				}
				for (String strtitleYear : setObj.envtitleYear) {

					try {
						titleYear = D8
								.findElementByXPath("//span[contains(@class,'"
										+ strtitleYear + "')]");
						titleYearnum = Integer.parseInt(titleYear.getText());

						usedTitleYear = strtitleYear;
						usedTitletag = "span";
						objectfound = 1;
						break;
					} catch (Exception e) {
						try {
							titleYear = D8
									.findElementByXPath("//select[contains(@class,'"
											+ strtitleYear + "')]");
							titleYearnum = Integer.parseInt(new Select(
									titleYear).getFirstSelectedOption()
									.getText());
							usedTitleYear = strtitleYear;
							usedTitletag = "select";
							objectfound = 1;

							break;
						} catch (Exception e5) {

						}
						continue;
					}
				}
				if (objectfound == 0) {
					update_Report("titleyear");
					break;
				} else {
					objectfound = 0;

				}

				switch (titleMonthString.toLowerCase()) {

				case "jan":
					monthnum1 = 1;
					break;
				case "1月":
					monthnum1 = 1;
					break;
				case "january":
					monthnum1 = 1;
					break;
				case "feb":
					monthnum1 = 2;
					break;
				case "2月":
					monthnum1 = 2;
					break;
				case "february":
					monthnum1 = 2;
					break;
				case "mar":
					monthnum1 = 3;
					break;
				case "3月":
					monthnum1 = 3;
					break;
				case "march":
					monthnum1 = 3;
					break;
				case "apr":
					monthnum1 = 4;
					break;
				case "4月":
					monthnum1 = 4;
					break;
				case "april":
					monthnum1 = 4;
					break;
				case "may":
					monthnum1 = 5;
					break;
				case "5月":
					monthnum1 = 5;
					break;
				case "june":
					monthnum1 = 6;
					break;
				case "jun":
					monthnum1 = 6;
					break;
				case "6月":
					monthnum1 = 6;
					break;
				case "july":
					monthnum1 = 7;
					break;
				case "jul":
					monthnum1 = 7;
					break;
				case "7月":
					monthnum1 = 7;
					break;
				case "aug":
					monthnum1 = 8;
					break;
				case "august":
					monthnum1 = 8;
					break;
				case "8月":
					monthnum1 = 8;
					break;
				case "sep":
					monthnum1 = 9;
					break;
				case "september":
					monthnum1 = 9;
					break;
				case "9月":
					monthnum1 = 9;
					break;
				case "oct":
					monthnum1 = 10;
					break;
				case "october":
					monthnum1 = 10;
					break;
				case "10月":
					monthnum1 = 10;
					break;
				case "nov":
					monthnum1 = 11;
					break;
				case "november":
					monthnum1 = 11;
					break;
				case "11月":
					monthnum1 = 11;
					break;
				case "dec":
					monthnum1 = 12;
					break;
				case "december":
					monthnum1 = 12;
					break;
				case "12月":
					monthnum1 = 12;
					break;
				default:
					update_Report("monthnotidentified");
					break;
				}

				if (setyyyy > titleYearnum && setmm > monthnum1) {
					try {
						do {

							nextMonth.click();
							monthnum1++;
							if (monthnum1 == 13)
								monthnum1 = 1;
							nextMonth = D8
									.findElementByXPath("//a[contains(@class,'"
											+ locatorNext + "')]");

							titleMonth = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleMonth + "')]");

							titleYear = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleYear + "')]");
							if (usedTitletag == "span") {
								titleMonthString = titleMonth.getText();
								titleYearnum = Integer.parseInt(titleYear
										.getText());
							} else if (usedTitletag == "select") {

								titleMonthString = new Select(titleMonth)
										.getFirstSelectedOption().getText()
										.toString();
								titleYearnum = Integer.parseInt(new Select(
										titleYear).getFirstSelectedOption()
										.getText());
							}

						} while (!((titleMonthString.toLowerCase().contains(
								monthpart.toLowerCase()) || titleMonthString
								.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
					} catch (Exception e) {
						isDateNotSelected = 1;
						LOGGER.log(Level.SEVERE, "IOException Occured in jquery case [date]- " +e.getMessage()); 
					}

				} else if (setyyyy < titleYearnum && setmm < monthnum1) {

					try {
						do {

							prevMonth.click();
							monthnum1--;
							if (monthnum1 == 0)
								monthnum1 = 12;
							prevMonth = D8
									.findElementByXPath("//a[contains(@class,'"
											+ locatorPrev + "')]");

							titleMonth = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleMonth + "')]");

							titleYear = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleYear + "')]");
							if (usedTitletag == "span") {
								titleMonthString = titleMonth.getText();
								titleYearnum = Integer.parseInt(titleYear
										.getText());
							} else if (usedTitletag == "select") {

								titleMonthString = new Select(titleMonth)
										.getFirstSelectedOption().getText()
										.toString();
								titleYearnum = Integer.parseInt(new Select(
										titleYear).getFirstSelectedOption()
										.getText());
							}

						} while (!((titleMonthString.toLowerCase().contains(
								monthpart.toLowerCase()) || titleMonthString
								.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
					} catch (Exception e) {
						isDateNotSelected = 1;
						LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date]- " +e.getMessage()); 
					}
				}

				else if (setyyyy == titleYearnum && setmm < monthnum1) {
					try {
						do {
							prevMonth.click();
							monthnum1--;
							if (monthnum1 == 0)
								monthnum1 = 12;
							prevMonth = D8
									.findElementByXPath("//a[contains(@class,'"
											+ locatorPrev + "')]");
							titleMonth = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleMonth + "')]");
							titleYear = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleYear + "')]");
							if (usedTitletag == "span") {
								titleMonthString = titleMonth.getText();
								titleYearnum = Integer.parseInt(titleYear
										.getText());
							} else if (usedTitletag == "select") {

								titleMonthString = new Select(titleMonth)
										.getFirstSelectedOption().getText()
										.toString();
								titleYearnum = Integer.parseInt(new Select(
										titleYear).getFirstSelectedOption()
										.getText());
							}
						} while (!((titleMonthString.toLowerCase().contains(
								monthpart.toLowerCase()) || titleMonthString
								.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
					} catch (Exception e) {
						isDateNotSelected = 1;
						LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date]- " +e.getMessage()); 
					}
				} else if (setyyyy == titleYearnum && setmm > monthnum1) {
					try {
						do {
							nextMonth.click();

							monthnum1++;
							if (monthnum1 == 13)
								monthnum1 = 1;
							nextMonth = D8

							.findElementByXPath("//a[contains(@class,'"
									+ locatorNext + "')]");

							titleMonth = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleMonth + "')]");

							titleYear = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleYear + "')]");
							if (usedTitletag == "span") {
								titleMonthString = titleMonth.getText();
								titleYearnum = Integer.parseInt(titleYear
										.getText());
							} else if (usedTitletag == "select") {

								titleMonthString = new Select(titleMonth)
										.getFirstSelectedOption().getText()
										.toString();
								titleYearnum = Integer.parseInt(new Select(
										titleYear).getFirstSelectedOption()
										.getText());
							}

						} while (!((titleMonthString.toLowerCase().contains(
								monthpart.toLowerCase()) || titleMonthString
								.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
					} catch (Exception e) {
						isDateNotSelected = 1;
						LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date]- " +e.getMessage()); 
					}
				}

				else if (setyyyy > titleYearnum && setmm == monthnum1) {
					try {
						do {

							nextMonth.click();
							monthnum1++;
							if (monthnum1 == 13)
								monthnum1 = 1;
							nextMonth = D8
									.findElementByXPath("//a[contains(@class,'"
											+ locatorNext + "')]");

							titleMonth = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleMonth + "')]");

							titleYear = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleYear + "')]");
							if (usedTitletag == "span") {
								titleMonthString = titleMonth.getText();
								titleYearnum = Integer.parseInt(titleYear
										.getText());
							} else if (usedTitletag == "select") {

								titleMonthString = new Select(titleMonth)
										.getFirstSelectedOption().getText()
										.toString();
								titleYearnum = Integer.parseInt(new Select(
										titleYear).getFirstSelectedOption()
										.getText());
							}

						} while (!((titleMonthString.toLowerCase().contains(
								monthpart.toLowerCase()) || titleMonthString
								.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
					} catch (Exception e) {
						isDateNotSelected = 1;
						LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date Calendar]- " +e.getMessage()); 
					}
				} else if (setyyyy < titleYearnum && setmm == monthnum1) {
					try {
						do {
							prevMonth.click();
							monthnum1--;
							if (monthnum1 == 0)
								monthnum1 = 12;
							prevMonth = D8
									.findElementByXPath("//a[contains(@class,'"
											+ locatorPrev + "')]");
							titleMonth = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleMonth + "')]");
							titleYear = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleYear + "')]");
							if (usedTitletag == "span") {
								titleMonthString = titleMonth.getText();
								titleYearnum = Integer.parseInt(titleYear
										.getText());
							} else if (usedTitletag == "select") {

								titleMonthString = new Select(titleMonth)
										.getFirstSelectedOption().getText()
										.toString();

								titleYearnum = Integer.parseInt(new Select(
										titleYear).getFirstSelectedOption()
										.getText());
							}

						} while (!((titleMonthString.toLowerCase().contains(
								monthpart.toLowerCase()) || titleMonthString
								.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
					} catch (Exception e) {
						isDateNotSelected = 1;
						LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date Calendar]- " +e.getMessage()); 
					}

				} else if (setyyyy > titleYearnum && setmm < monthnum1) {
					try {
						do {

							nextMonth.click();
							monthnum1++;
							if (monthnum1 == 13)
								monthnum1 = 1;
							nextMonth = D8

							.findElementByXPath("//a[contains(@class,'"
									+ locatorNext + "')]");

							titleMonth = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleMonth + "')]");

							titleYear = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleYear + "')]");
							if (usedTitletag == "span") {
								titleMonthString = titleMonth.getText();
								titleYearnum = Integer.parseInt(titleYear
										.getText());
							} else if (usedTitletag == "select") {

								titleMonthString = new Select(titleMonth)
										.getFirstSelectedOption().getText()
										.toString();
								titleYearnum = Integer.parseInt(new Select(
										titleYear).getFirstSelectedOption()
										.getText());
							}
						} while (!((titleMonthString.toLowerCase().contains(
								monthpart.toLowerCase()) || titleMonthString
								.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
					} catch (Exception e) {
						isDateNotSelected = 1;
						LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date Calendar]- " +e.getMessage()); 
					}
				} else if (setyyyy < titleYearnum && setmm > monthnum1) {
					try {
						do {
							prevMonth.click();
							monthnum1--;
							if (monthnum1 == 0)
								monthnum1 = 12;
							prevMonth = D8
									.findElementByXPath("//a[contains(@class,'"
											+ locatorPrev + "')]");
							titleMonth = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleMonth + "')]");
							titleYear = D8.findElementByXPath("//"
									+ usedTitletag + "[contains(@class,'"
									+ usedTitleYear + "')]");
							if (usedTitletag == "span") {
								titleMonthString = titleMonth.getText();
								titleYearnum = Integer.parseInt(titleYear
										.getText());
							} else if (usedTitletag == "select") {

								titleMonthString = new Select(titleMonth)
										.getFirstSelectedOption().getText()
										.toString();
								titleYearnum = Integer.parseInt(new Select(
										titleYear).getFirstSelectedOption()
										.getText());
							}
						} while (!((titleMonthString.toLowerCase().contains(
								monthpart.toLowerCase()) || titleMonthString
								.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
					} catch (Exception e) {
						isDateNotSelected = 1;
						LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date Calendar]- " +e.getMessage()); 
					}
				}
				if (isDateNotSelected == 0) {
					if (setyyyy == titleYearnum && setmm == monthnum1) {

						try {

							List<WebElement> daystoClick = D8
									.findElementsByXPath("//a[contains(@class,'"
											+ locatorPrev
											+ "')]"
											+ "/following::a[contains(text(),'"
											+ setdd + "')]");

							List<WebElement> categoryDeciders = D8
									.findElementsByXPath("//a[contains(@class,'"
											+ locatorPrev
											+ "')]"
											+ "/following::a[contains(text(),'"
											+ 15 + "')]");

							for (WebElement categoryDecider : categoryDeciders) {
								if (categoryDecider.getAttribute("href")
										.endsWith("#")
										&& categoryDecider.getText().equals(
												"15")) {
									realCategoryDeciders.add(categoryDecider);
									dateClass = categoryDecider
											.getAttribute("class");

								}
							}

							for (int i2 = 0; i2 < daystoClick.size(); i2++) {
								
								if (today == setdd
										&& Integer.parseInt(daystoClick.get(i2)
												.getText()) == setdd
										&& daystoClick.get(i2)
												.getAttribute("href")
												.endsWith("#")) {

									daytoClick = daystoClick.get(i2);
									daytoClick.click();
									lastSelecteddateJQ = setdd;
									lastSelectedmonthJQ = setmm;
									lastSelectedyearJQ = setyyyy;
									isDateSelected = 1;
									update_Report("executed", ObjectSet,
											ObjectSetVal);
									break;

								} else if (Integer.parseInt(daystoClick.get(i2)
										.getText()) == setdd
										&& (daystoClick.get(i2)
												.getAttribute("class")
												.equalsIgnoreCase(dateClass) || (daystoClick
												.get(i2).getAttribute("class")
												.equalsIgnoreCase(selectedDateClass))
												&& daystoClick.get(i2)
														.getAttribute("href")
														.endsWith("#"))) {

									daytoClick = daystoClick.get(i2);
									daytoClick.click();
									lastSelecteddateJQ = setdd;
									lastSelectedmonthJQ = setmm;
									lastSelectedyearJQ = setyyyy;
									isDateSelected = 1;
									selectedDateClass = daytoClick
											.getAttribute("class");
									update_Report("executed", ObjectSet,
											ObjectSetVal);
									break;
								}
							}

						} catch (Exception e1) {
							isDateNotSelected = 1;
							LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date Calendar]- " +e1.getMessage()); 

						}

					}
				 
					if (captureperform == true) {
						screenShot(loopnum, TScrowcount, TScname);
					}
					if (isDateSelected == 0) {

						update_Report("invaliddate");
						robot1.keyPress(KeyEvent.VK_ESCAPE);
						robot1.keyRelease(KeyEvent.VK_ESCAPE);
					} else {
						isDateSelected = 0;
					}

				} else {

					isDateNotSelected = 0;
					update_Report("invaliddate");
					robot1.keyPress(KeyEvent.VK_ESCAPE);
					robot1.keyRelease(KeyEvent.VK_ESCAPE);

				}
			} else {

				update_Report("executed", ObjectSet, ObjectSetVal);
				robot1.keyPress(KeyEvent.VK_ESCAPE);
				robot1.keyRelease(KeyEvent.VK_ESCAPE);
				if (captureperform == true) {
					screenShot(loopnum, TScrowcount, TScname);
				}
			}
 
			break;

		case "boot":

			for (String strPrevMonth : setObj.envprevMonth1) {
				try {
					prevMonth = D8.findElementByXPath("//th[contains(@class,'"
							+ strPrevMonth + "')]");
					locatorPrev = strPrevMonth;
					objectfound = 1;
					break;
				} catch (Exception e) {
					continue;
				}
			}
			if (objectfound == 0) {
				update_Report("prevmonth");
				break;
			} else {
				objectfound = 0;
			}

			for (String strNextMonth : setObj.envnextMonth1) {
				try {
					nextMonth = D8.findElementByXPath("//th[contains(@class,'"
							+ strNextMonth + "')]");
					locatorNext = strNextMonth;
					objectfound = 1;
					break;
				} catch (Exception e) {
					continue;
				}
			}
			if (objectfound == 0) {
				update_Report("nextmonth");
				break;
			} else {
				objectfound = 0;
			}

			try {
				titleDefault = D8.findElementByXPath("//th[contains(@class,'"
						+ "switch" + "')]");

			}

			catch (Exception e) {
				update_Report("titledefault");
				break;
			}
			titleDefaultText = titleDefault.getText().trim();
			titleLength = titleDefaultText.length();
			titleYearnum = Integer.parseInt(titleDefaultText
					.substring(titleLength - 4));
			titleMonthString = titleDefaultText.substring(0, titleLength - 5)
					.trim();
			switch (titleMonthString.toLowerCase()) {

			case "jan":
				monthnum1 = 1;
				break;
			case "1月":
				monthnum1 = 1;
				break;
			case "january":
				monthnum1 = 1;
				break;
			case "feb":
				monthnum1 = 2;
				break;
			case "2月":
				monthnum1 = 2;
				break;
			case "february":
				monthnum1 = 2;
				break;
			case "mar":
				monthnum1 = 3;
				break;
			case "3月":
				monthnum1 = 3;
				break;
			case "march":
				monthnum1 = 3;
				break;
			case "apr":
				monthnum1 = 4;
				break;
			case "4月":
				monthnum1 = 4;
				break;
			case "april":
				monthnum1 = 4;
				break;
			case "may":
				monthnum1 = 5;
				break;
			case "5月":
				monthnum1 = 5;
				break;
			case "june":
				monthnum1 = 6;
				break;
			case "jun":
				monthnum1 = 6;
				break;
			case "6月":
				monthnum1 = 6;
				break;
			case "july":
				monthnum1 = 7;
				break;
			case "jul":
				monthnum1 = 7;
				break;
			case "7月":
				monthnum1 = 7;
				break;
			case "aug":
				monthnum1 = 8;
				break;
			case "august":
				monthnum1 = 8;
				break;
			case "8月":
				monthnum1 = 8;
				break;
			case "sep":
				monthnum1 = 9;
				break;
			case "september":
				monthnum1 = 9;
				break;
			case "9月":
				monthnum1 = 9;
				break;
			case "oct":
				monthnum1 = 10;
				break;
			case "october":
				monthnum1 = 10;
				break;
			case "10月":
				monthnum1 = 10;
				break;
			case "nov":
				monthnum1 = 11;
				break;
			case "november":
				monthnum1 = 11;
				break;
			case "11月":
				monthnum1 = 11;
				break;
			case "dec":
				monthnum1 = 12;
				break;
			case "december":
				monthnum1 = 12;
				break;
			case "12月":
				monthnum1 = 12;
				break;
			default:
				update_Report("monthnotidentified");
				break;
			}
			try {
				nextMonth = D8.findElementByXPath("//th[contains(@class,'"
						+ locatorNext + "')]");
			} catch (Exception e) {
				update_Report("nextmonth");
				break;
			}
			try {
				prevMonth = D8.findElementByXPath("//th[contains(@class,'"
						+ locatorPrev + "')]");
			} catch (Exception e) {
				update_Report("prevmonth");
				LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [Switch case of Month ]- " +e.getMessage()); 
				break;
			}
			if (setyyyy > titleYearnum && setmm > monthnum1) {
				try {
					do {

						nextMonth.click();
						monthnum1++;
						if (monthnum1 == 13)
							monthnum1 = 1;
						titleDefault = D8
								.findElementByXPath("//th[contains(@class,'"
										+ "switch" + "')]");
						titleDefaultText = titleDefault.getText().trim();
						titleLength = titleDefaultText.length();
						titleYearnum = Integer.parseInt(titleDefaultText
								.substring(titleLength - 4));
						titleMonthString = titleDefaultText.substring(0,
								titleLength - 5).trim();

					}

					while (!((titleMonthString.toLowerCase().contains(
							monthpart.toLowerCase()) || titleMonthString
							.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
				} catch (Exception e) {
					update_Report("invaliddate");
					LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date Calendar]- " +e.getMessage()); 
					break;
				}
			} else if (setyyyy < titleYearnum && setmm < monthnum1) {
				try {
					do {

						prevMonth.click();
						monthnum1--;
						if (monthnum1 == 0)
							monthnum1 = 12;
						titleDefault = D8
								.findElementByXPath("//th[contains(@class,'"
										+ "switch" + "')]");
						titleDefaultText = titleDefault.getText().trim();
						titleLength = titleDefaultText.length();
						titleYearnum = Integer.parseInt(titleDefaultText
								.substring(titleLength - 4));
						titleMonthString = titleDefaultText.substring(0,
								titleLength - 5).trim();
					} while (!((titleMonthString.toLowerCase().contains(
							monthpart.toLowerCase()) || titleMonthString
							.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
				} catch (Exception e) {
					update_Report("invaliddate");
					LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date invaliddate]- " +e.getMessage()); 
					break;
				}
			}

			else if (setyyyy == titleYearnum && setmm < monthnum1) {
				try {
					do {
						prevMonth.click();
						monthnum1--;
						if (monthnum1 == 0)
							monthnum1 = 12;
						titleDefault = D8
								.findElementByXPath("//th[contains(@class,'"
										+ "switch" + "')]");
						titleDefaultText = titleDefault.getText().trim();
						titleLength = titleDefaultText.length();
						titleYearnum = Integer.parseInt(titleDefaultText
								.substring(titleLength - 4));
						titleMonthString = titleDefaultText.substring(0,
								titleLength - 5).trim();
					} while (!((titleMonthString.toLowerCase().contains(
							monthpart.toLowerCase()) || titleMonthString
							.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
				} catch (Exception e) {
					update_Report("invaliddate");
					LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date invaliddate]- " +e.getMessage()); 
					break;
				}
			} else if (setyyyy == titleYearnum && setmm > monthnum1) {
				try {
					do {
						nextMonth.click();
						monthnum1++;
						if (monthnum1 == 13)
							monthnum1 = 1;
						titleDefault = D8
								.findElementByXPath("//th[contains(@class,'"
										+ "switch" + "')]");
						titleDefaultText = titleDefault.getText().trim();
						titleLength = titleDefaultText.length();
						titleYearnum = Integer.parseInt(titleDefaultText
								.substring(titleLength - 4));
						titleMonthString = titleDefaultText.substring(0,
								titleLength - 5).trim();
					} while (!((titleMonthString.toLowerCase().contains(
							monthpart.toLowerCase()) || titleMonthString
							.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
				} catch (Exception e) {
					update_Report("invaliddate");
					LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date invaliddate]- " +e.getMessage()); 
					break;
				}

			}

			else if (setyyyy > titleYearnum && setmm == monthnum1) {
				try {
					do {
						nextMonth.click();
						monthnum1++;
						if (monthnum1 == 13)
							monthnum1 = 1;
						titleDefault = D8
								.findElementByXPath("//th[contains(@class,'"
										+ "switch" + "')]");
						titleDefaultText = titleDefault.getText().trim();
						titleLength = titleDefaultText.length();
						titleYearnum = Integer.parseInt(titleDefaultText
								.substring(titleLength - 4));
						titleMonthString = titleDefaultText.substring(0,
								titleLength - 5).trim();
					} while (!((titleMonthString.toLowerCase().contains(
							monthpart.toLowerCase()) || titleMonthString
							.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
				} catch (Exception e) {
					update_Report("invaliddate");
					LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date invaliddate]- " +e.getMessage()); 
					break;
				}
			} else if (setyyyy < titleYearnum && setmm == monthnum1) {
				try {
					do {
						prevMonth.click();
						monthnum1--;
						if (monthnum1 == 0)
							monthnum1 = 12;
						titleDefault = D8
								.findElementByXPath("//th[contains(@class,'"
										+ "switch" + "')]");
						titleDefaultText = titleDefault.getText().trim();
						titleLength = titleDefaultText.length();
						titleYearnum = Integer.parseInt(titleDefaultText
								.substring(titleLength - 4));
						titleMonthString = titleDefaultText.substring(0,
								titleLength - 5).trim();

					} while (!((titleMonthString.toLowerCase().contains(
							monthpart.toLowerCase()) || titleMonthString
							.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
				} catch (Exception e) {
					update_Report("invaliddate");
					LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date invaliddate]- " +e.getMessage()); 
					break;
				}
			} else if (setyyyy > titleYearnum && setmm < monthnum1) {
				try {
					do {
						nextMonth.click();
						monthnum1++;
						if (monthnum1 == 13)
							monthnum1 = 1;
						titleDefault = D8
								.findElementByXPath("//th[contains(@class,'"
										+ "switch" + "')]");
						titleDefaultText = titleDefault.getText().trim();
						titleLength = titleDefaultText.length();
						titleYearnum = Integer.parseInt(titleDefaultText
								.substring(titleLength - 4));
						titleMonthString = titleDefaultText.substring(0,
								titleLength - 5).trim();
					} while (!((titleMonthString.toLowerCase().contains(
							monthpart.toLowerCase()) || titleMonthString
							.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
				} catch (Exception e) {
					update_Report("invaliddate");
					LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date invaliddate]- " +e.getMessage()); 
					break;
				}
			} else if (setyyyy < titleYearnum && setmm > monthnum1) {
				try {
					do {
						prevMonth.click();
						monthnum1--;
						if (monthnum1 == 0)
							monthnum1 = 12;
						titleDefault = D8
								.findElementByXPath("//th[contains(@class,'"
										+ "switch" + "')]");
						titleDefaultText = titleDefault.getText().trim();
						titleLength = titleDefaultText.length();
						titleYearnum = Integer.parseInt(titleDefaultText
								.substring(titleLength - 4));
						titleMonthString = titleDefaultText.substring(0,
								titleLength - 5).trim();
					} while (!((titleMonthString.toLowerCase().contains(
							monthpart.toLowerCase()) || titleMonthString
							.toLowerCase().contains(monthpartjap)) && (setyyyy == titleYearnum)));
				} catch (Exception e) {
					update_Report("invaliddate");
					LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date invaliddate]- " +e.getMessage()); 
					break;
				}
			}
			if (setyyyy == titleYearnum && setmm == monthnum1) {
				try {
					List<WebElement> daystoClick = D8
							.findElementsByXPath("//td[contains(text(),'"
									+ setdd + "')]");

					if (daystoClick.size() == 1) {
						daytoClick = daystoClick.get(0);
						daytoClick.click();
						update_Report("executed", ObjectSet, ObjectSetVal);
					} else if (daystoClick.size() > 1) {
						if (setdd <= 7) {
							for (int i2 = 0; i2 < daystoClick.size(); i2++) {

								if (Integer.parseInt(daystoClick.get(i2)
										.getText()) == setdd) {

									daytoClick = daystoClick.get(i2);
									daytoClick.click();
									update_Report("executed", ObjectSet,
											ObjectSetVal);
									break;
								}
							}
						 

						} else {
							for (int i2 = daystoClick.size() - 1; i2 >= 0; i2--) {

								if (Integer.parseInt(daystoClick.get(i2)
										.getText()) == setdd) {

									daytoClick = daystoClick.get(i2);
									daytoClick.click();
									update_Report("executed", ObjectSet,
											ObjectSetVal);
									break;
								}
							}

						}
					} else {
						update_Report("invaliddate");
						LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date invaliddate]- " ); 
						break;
					}

				} catch (Exception ex1) {
					try {
						List<WebElement> daystoClick = D8
								.findElementsByXPath("//span[contains(text(),'"
										+ setdd + "')]");
						if (daystoClick.size() == 1) {
							daytoClick = daystoClick.get(0);
							daytoClick.click();
							update_Report("executed", ObjectSet, ObjectSetVal);
						} else if (daystoClick.size() > 1) {
							if (setdd <= 7) {
								for (int i2 = 0; i2 < daystoClick.size(); i2++) {

									if (Integer.parseInt(daystoClick.get(i2)
											.getText()) == setdd) {

										daytoClick = daystoClick.get(i2);
										daytoClick.click();
										update_Report("executed", ObjectSet,
												ObjectSetVal);
										break;
									}
								}
								

							} else {
								for (int i2 = daystoClick.size() - 1; i2 >= 0; i2--) {

									if (Integer.parseInt(daystoClick.get(i2)
											.getText()) == setdd) {

										daytoClick = daystoClick.get(i2);
										daytoClick.click();
										update_Report("executed", ObjectSet,
												ObjectSetVal);
										break;
									}
								}

							}
						} else {
							update_Report("invaliddate");
						}

						update_Report("executed", ObjectSet, ObjectSetVal);
					}

					catch (Exception e) {
						update_Report("invaliddate");
						LOGGER.log(Level.SEVERE, " Exception Occured in jquery case [date invaliddate]- " +e.getMessage()); 
					}
				}

			}
			if (captureperform == true) {
				screenShot(loopnum, TScrowcount, TScname);
			}
			elem.click();

			break;

		default:
			System.out.println("Notdefined");
			break;

		}

	}
	
	public static void func_InvokeFunction(String functionName,
			String argumentlist) {
		Object[] argument_list = null;
		int checkNULL = 0;
		int CheckONE = 0;
		Class[] parameterTypes = null;
		// 2016-03-23 callfunction no argument Support
	 
		if (argumentlist.isEmpty()) {
			checkNULL = 1;

		} else if (argumentlist.contains(":")) {
			argument_list = argumentlist.split(":");
		} else {
			CheckONE = 1;

		}
		String function_name = functionName;
		try {
			SeleniumWD s1 = new SeleniumWD();
			Method[] declaredMethods = SeleniumWD.class.getDeclaredMethods();
			for (Method m : declaredMethods) {
				System.out.println(m.getName());
				if (checkNULL != 1) {
					parameterTypes = m.getParameterTypes();
				}
				if (checkNULL == 0) {
					if ((m.getName()).equals(function_name)) {
						if (parameterTypes.length > 1) {
							if (parameterTypes.length == argument_list.length) {
								try {
									m.invoke(s1, argument_list);
									update_Report("executed");
								} catch (Exception e) {
									update_Report("userdefined");
								}
							}
							break;
						} else if ((m.getName()).equals(function_name)
								&& CheckONE == 1 && parameterTypes.length == 1) {

							try {
								m.invoke(s1, argumentlist);
								update_Report("executed");
							} catch (Exception e) {
								update_Report("userdefined");
							}

							break;
						}
					}
				} else if (m.getName().equals(function_name) && checkNULL == 1) {
					try {
						m.invoke(s1);
						update_Report("executed");
					} catch (Exception e) {
						update_Report("userdefined");
					}
					break;
				}

			}

		} catch (Exception e) {
			System.out.println(e);
			LOGGER.log(Level.SEVERE, " Exception Occured in func_InvokeFunction- " +e.getMessage()); 
		}
	}
	
	public static boolean isInteger(String input) {
		try {
			Integer.parseInt(input);
			return true;
		} catch (Exception e) {
			LOGGER.log(Level.SEVERE, " Exception Occured in isInteger- " +e.getMessage()); 
			return false;
		}
	}

//--- 2015-06-15 sample method for "callfunction" --- 
	// Single Argument
	public static void Display(String Name)
	{

				System.out.println(Name);
	}
// three
	public static void Display(String Name,String Name2,String Name3)
	{

				System.out.println(Name+ " "+Name2+ " "+Name3+ " ");
	}
		
	public static void CalculationTest(String value1, String value2, String value3, String value4)
	{
		int convert_value1 = Integer.parseInt(value1);
		int convert_value2 = Integer.parseInt(value2);
		int convert_value3 = Integer.parseInt(value3);
		int convert_value4 = Integer.parseInt(value4);
		int Add = 0;
		int Subtraction = 0;
		int Multiplication = 0;
		int Division = 0;

		Add = convert_value1 + convert_value2;
		Subtraction = convert_value3 - convert_value4;
		Multiplication = convert_value1 * convert_value3;
		Division = convert_value2 / convert_value4;

		System.out.println("value1=" + value1);
		System.out.println("value2=" + value2);
		System.out.println("value3=" + value3);
		System.out.println("value4=" + value4);
		System.out.println("Add = value1 + value2");
		System.out.println("Add="+ Add);
		System.out.println("Subtraction = value3 - value4");
		System.out.println("Subtraction=" + Subtraction);
		System.out.println("Multiplication = value1 * value3");
		System.out.println("Multiplication=" + Multiplication);
		System.out.println("Division = value2 / value4");
		System.out.println("Division=" + Division);
		
	}
}
