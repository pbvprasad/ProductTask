package com.test.testResources;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.net.InetAddress;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.Reporter;
import org.testng.asserts.SoftAssert;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.Font;
import com.itextpdf.text.Font.FontFamily;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfWriter;
import com.test.etochaTestProject.EtochaTest;

 @SuppressWarnings("unused")
public class commnFunctions extends pageObjects {

	//function to wait/ sync for given amount of time in seconds==================================================================s==============================================
	public void fnAppSync(int intTime) throws Exception {
		Robot robot = new Robot();
		robot.delay(intTime * 500); }
 
	//Function to get the given object exists or not================================================================================================================================
	public int fnObjectExists(String strObject, int timeInterval) throws Throwable {
		int strDisp = 0;
		for(int intVal=0; intVal <= timeInterval; intVal++) {
			try {
				if (wDriver.findElements(By.xpath(strObject)).size() > 0)
					if (wDriver.findElement(By.xpath(strObject)).isDisplayed() == true) {
								strDisp = 1; 
								fnObjectHighLight(strObject); } } 
			catch (Exception exp) {
				log.info(exp.getMessage());
				strDisp = 0; 
				Thread.sleep(100); }
			if (strDisp == 1) break; }
		
		if (strDisp==1) log.info(" Object ID - " + strObject + " No. of Objects Found : " + strDisp);
		else log.error(" Object ID - " + strObject + " No. of Objects Found : " + strDisp);

		return strDisp; }
	
	public void fnObjectHighLight(String strObj) throws Throwable {
		WebElement ele = wDriver.findElement(By.xpath(strObj));		        
		JavascriptExecutor js = (JavascriptExecutor) wDriver;
		js.executeScript("arguments[0].setAttribute('style', 'background: ; border: 1px solid red;');", ele);
		Thread.sleep(75);
		js.executeScript("arguments[0].setAttribute('style', 'background: ; border: ;');", ele); }
	
	//Method to close all the chrome browsers if any opened ===============================================
	public void fnCloseBrowsers() {
		try{
			Runtime.getRuntime().exec("taskkill /IM chrome.exe /f");
			fnAppSync(2);
			Runtime.getRuntime().exec("taskkill /IM chromedriver.exe /f");
			}
		catch (Exception exp) {
			log.error("[INFO]    Unable to close/kill the chrome process from the task manager"); } }

	//Method to select the specified product from the menu ========================================================================
	public boolean fnValidatePageNavigationStatus(String objClickObj, String objClickObjLabel, String objCheckObj) throws Throwable {
		boolean strRetVal = false;

		if (fnObjectExists(objClickObj, 5)!=0) {
			fnObjectClick(objClickObj);
			fnAppSync(5);
			if (fnObjectExists(objCheckObj, 5)!=0) {
				log.info("Successfully selected/clicked on the object " + objCheckObj);
				log.info("Successfully navigated");
				Reporter.log("Successfully selected/clicked on the object " + objClickObjLabel);
				strRetVal = true;
				fnCaptureScreenShot("Successfully selected/clicked on the object " + objClickObjLabel); }
			else {
				log.error("Failed to select/click on the object " + objCheckObj);
				Reporter.log("Failed to select/click on the object " + objClickObjLabel);
				testAssert.fail("Failed to select/click on the object " + objClickObjLabel);
				stepStatus = false;	
				fnCaptureScreenShot("Failed to select/click on the object " + objClickObjLabel); } }
		
		return strRetVal;	}

	//Function to click on the element=================================================================
	public void fnObjectClick(String strObj) throws Throwable{
		if (fnObjectExists(strObj, 1) != 0) {
			try{ wDriver.findElement(By.xpath(strObj)).click();	
				log.info("[INFO]    1. Object Click successful - " + strObj);} 
			catch (Exception exp) {
				log.info("[INFO]    2. Trying to click on - " + strObj);
				fnAppSync(50);
				try {
					wDriver.findElement(By.xpath(strObj)).click(); }
				catch (Exception exp1) {
					log.error("[ERROR]    Unable to click on the object : " + strObj);	} } } }
	
	//Function to launch the chrome browser ======================================================
	public void dpmChromeBrowserLaunch(String dataSheet, String dataPointer) throws Throwable {
		String appUrl = fnGetExcelData(pageObjects.strDefaultPath, dataSheet, dataPointer, "LoginURL");
		pageObjects.strEnvironmentURL = appUrl;
		try {
			//Step to kill all the previously opened browsers if any
			fnCloseBrowsers();
			
			log.info("Launching the chrome browser - 1st time");
			log.info("Chrome Driver path - " + pageObjects.strChromeDriverPathName);
			wDriver = null;
			
			File file = new File(pageObjects.strChromeDriverPathName);
			System.setProperty("webdriver.chrome.driver", file.getAbsolutePath());
			ChromeOptions options = new ChromeOptions();
			options.addArguments(("--disable-notifications"));
			options.addArguments("disable-infobars");
			options.addArguments("--lang=es");
			options.addArguments("--start-maximized");
			options.addArguments("window-size=1024,768");
			Map<String, Object> prefs = new HashMap<String, Object>();
		    prefs.put("credentials_enable_service", false);
		    prefs.put("profile.password_manager_enabled", false);
		    options.setExperimentalOption("prefs", prefs);
			wDriver = new ChromeDriver(options);
			wDriver.get(appUrl);
			wDriver.manage().window().maximize();
			wDriver.manage().deleteAllCookies(); 
			wDriver.manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
			wDriver.manage().timeouts().setScriptTimeout(30, TimeUnit.SECONDS);
			wDriver.manage().timeouts().implicitlyWait(1, TimeUnit.SECONDS); 
			
			log.info("Successfully opened the chrome browser");
			Reporter.log("Successfully opened the chrome browser");	}
		catch (Exception exp) {
			log.error("Failed to open the chrome browser");
			Reporter.log("Failed to open the chrome browser");
			log.info(exp.getStackTrace()); 
			pageObjects.stepStatus = false; 
			testAssert.fail("Failed to open the chrome browser");	} }

	@SuppressWarnings("deprecation")
	public synchronized String fnGetExcelData(String fileName, String sheetName, String dataPointer, String columnName) throws Throwable {
		String cellValue = "";
		//log.info("fnGetExcel Data --- " + fileName);
		//get the rows count from the excel sheet
		int rowNum = fnGetExcelRowNum(fileName, sheetName, dataPointer);		
		if (rowNum == -1){
			log.info("[WARNING]    Incorrect DataPointer - " + dataPointer); }
		
		//get the columns count from the excel sheet
		int colNum = fnGetExcelColNum(fileName, sheetName, columnName);
		if (colNum == -1){
			log.info("[WARNING]    Incorrect column name - " + columnName); }

		if (rowNum > -1 && colNum > -1) {
			//String strTemp = fnGetExcelCellData(fileName, sheetName, rowNum, columnName);
			File file=new File(fileName);
	        if(file.exists()) {
		        InputStream myxls = new FileInputStream(fileName); 
		        @SuppressWarnings("resource")
				HSSFWorkbook book = new HSSFWorkbook(myxls);
		        HSSFSheet sheet = book.getSheet(sheetName);
		        Row row = sheet.getRow(rowNum);
		        Cell cell1 = row.getCell(colNum);
     
		        if (cell1 == null)
		        	cellValue = "";
		        else if(cell1.getCellType() == Cell.CELL_TYPE_NUMERIC)
		        	cellValue = NumberToTextConverter.toText(cell1.getNumericCellValue());
		        else if(cell1.getCellType() == Cell.CELL_TYPE_STRING)
		        	cellValue = cell1.getStringCellValue();
		        else if(cell1.getCellType() == Cell.CELL_TYPE_BLANK)
		        	cellValue = "";
		        else if(cell1.getCellType() == Cell.CELL_TYPE_FORMULA){
		        	FormulaEvaluator evaluator = book.getCreationHelper().createFormulaEvaluator();
		        	evaluator.evaluateFormulaCell(cell1);
		        	if (cell1.getCachedFormulaResultType() == HSSFCell.CELL_TYPE_STRING)
		        		cellValue = cell1.getStringCellValue();
		        	else if (cell1.getCachedFormulaResultType() == HSSFCell.CELL_TYPE_NUMERIC)
		        		cellValue = NumberToTextConverter.toText(cell1.getNumericCellValue());
		        	else if(cell1.getCachedFormulaResultType() == Cell.CELL_TYPE_BLANK)
			        	cellValue = ""; }
		        myxls.close(); }
	        		        	
	        log.info("[INFO]   == Cell Row : " + rowNum + " Datapointer : " + dataPointer + " - Cell Column : " + colNum + ", Column Name : " + columnName + " - Cell Value : " + cellValue); }
	       else if ((rowNum == -1) || (colNum == -1)) {
	    	   cellValue = "";
	    	   //log.info("[INFO]    Incorrect excel column/ row values == Cell Row : " + rowNum + " - Cell Column : " + colNum + ", Column Name : " + columnName + " - Cell Value : " + cellValue); 
	    	   }
		return cellValue.trim(); }
		
	@SuppressWarnings("deprecation")
	public synchronized int fnGetExcelRowNum(String fileName, String sheetName, String dataPointer) throws Exception {
		File file=new File(fileName);
		int colNum = -1;
        if(file.exists()){
	        InputStream myxls = new FileInputStream(new File(fileName));
	        HSSFWorkbook book = new HSSFWorkbook(myxls);
	        HSSFSheet sheet = book.getSheet(sheetName);
	        Row row = sheet.getRow(pageObjects.intSheetRowNum);
	        int rCnt = sheet.getLastRowNum();
	        String strCellValue = "";
	        for (int intLoop=0; intLoop <= rCnt; intLoop++) {
	        	row = sheet.getRow(intLoop);
	        	Cell cell = row.getCell(0);
	        	if (cell == null)
	        		strCellValue = "";
		        else if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
	        		strCellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
	 	        else if(cell.getCellType() == Cell.CELL_TYPE_STRING)
	 	        	strCellValue = cell.getStringCellValue();
	 	       else if(cell.getCellType() == Cell.CELL_TYPE_BLANK)
	 	    	  strCellValue = "";
	        	
	        	//log.info("Cell Value - " + strCellValue);
	        	if (strCellValue.trim().toLowerCase().contentEquals(dataPointer.trim().toLowerCase()))
	        		colNum = intLoop;
	        	if (colNum != -1) break; } 
	    book.close(); }
		return colNum;	}
	
	@SuppressWarnings("deprecation")
	public synchronized int fnGetExcelColNum(String fileName, String sheetName, String colName) throws Exception {
		File file=new File(fileName);
		int colNum = -1;
        if(file.exists()){
	        InputStream myxls = new FileInputStream(fileName); 
	        HSSFWorkbook book = new HSSFWorkbook(myxls); 
	        HSSFSheet sheet = book.getSheet(sheetName);
	        Row row = sheet.getRow(pageObjects.intSheetRowNum);
	        String strCellValue="";
	        for (int intLoop=0; intLoop < sheet.getRow(pageObjects.intSheetRowNum).getPhysicalNumberOfCells(); intLoop++) {
	        	Cell cell = row.getCell(intLoop);
	        	if (cell == null)
	        		strCellValue = "";
		        else if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
	        		strCellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
	 	        else if(cell.getCellType() == Cell.CELL_TYPE_STRING)
	 	        	strCellValue = cell.getStringCellValue();
	 	       else if(cell.getCellType() == Cell.CELL_TYPE_BLANK)
	 	    	  strCellValue = "";
	        	if (cell.getStringCellValue().trim().toLowerCase().contentEquals(colName.trim().toLowerCase()))
	        		colNum = intLoop;
	        	if (colNum != -1) break; }
	        book.close();
	        myxls.close(); }    
		return colNum; }
	
	public void fnScreenCapture(WebDriver wDriver, String strFileName) throws Throwable {
		File scrFile;
		String fileName="";

		if (pageObjects.strScreenCapture == true) {
			log.info("[INFO]   taking screen capture in fnScreenCapture - " + strFileName);
	        String format = ".jpg";
	        if (pageObjects.strScrnFileName == false)
	        	fileName = strFileName;
	        else
	        	fileName = strFileName;
			try{ File file = new File(fileName);
				File myFile = new File(file, fileName);
				if (myFile.exists() == true)
					myFile.delete(); }  catch(Exception e) { 
						log.error("[ERROR]    Unable to delete the file ... " + fileName);
						Reporter.log("[ERROR]    Unable to delete the file ... " + fileName); }
			try{
				scrFile = ((TakesScreenshot)wDriver).getScreenshotAs(OutputType.FILE);
				FileUtils.copyFile(scrFile, new File(fileName));
			} 
			catch (Exception exp) {
				try {
				log.info("[INFO]   taking screen capture in catch statement - fnScreenCapture - " + strFileName);
				Reporter.log("[INFO]   taking screen capture in catch statement - fnScreenCapture - " + strFileName);
				scrFile = ((TakesScreenshot)wDriver).getScreenshotAs(OutputType.FILE); }
				catch (Exception exp1){
					pageObjects.strScnStatus = false;
					log.error("[INFO]   Failed to capture the screen, no browser exists - fnScreenCapture - " + strFileName);
					Reporter.log("[INFO]   Failed to capture the screen, no browser exists - fnScreenCapture - " + strFileName); } }	} }

	public void fnLoadConfigurationParameters() throws Throwable {
		pageObjects.stepStatus = true;
		PropertyConfigurator.configure(log4jConfigFile);
		log = Logger.getLogger(EtochaTest.class.getName());
		
        log.info("Successfully configured the log files");
        Reporter.log("Successfully configured the log files");
		
		InetAddress IP=InetAddress.getLocalHost();
		log.info("[INFO]   Running on machine - IP == " + IP.getHostAddress());
		Reporter.log("[INFO]   Running on machine - IP == " + IP.getHostAddress());
		pageObjects.strDataSheetPath = pageObjects.strDefaultPath; 
		fnCloseBrowsers();
		log.info("Successfully closed the existing chrome browsers if any");
		Reporter.log("Successfully closed the existing chrome browsers if any");
		testAssert = new SoftAssert(); }
	
	public synchronized static int fnGetExcelRowCount(String fileName, String sheetName) throws Exception {
		log.info("[INFO]   Data File Name : " + fileName);
		log.info("[INFO]   Data Sheet Name : " + sheetName);
		
		log.info("Data File Name : " + fileName);
		log.info("Data Sheet Name : " + sheetName);
		
		File file=new File(fileName);
		int rowCount = 0;
		try{
	        if(file.exists()){
		        InputStream myxls = new FileInputStream(fileName); 
		        HSSFWorkbook book = new HSSFWorkbook(myxls); 
		        HSSFSheet sheet = book.getSheet(sheetName);
		        rowCount = sheet.getLastRowNum();
		        book.close();
		        myxls.close(); } }
		catch (Exception exp){
			log.error("[ERROR]    File not found .... " + fileName); 
			testAssert.fail(exp.getMessage()); }
		return rowCount; }

	//Function to move the mouse pointer to the object =====================================================
	public void fnMouseMove(String strObj) throws Throwable {
		if (fnObjectExists(strObj, 1) != 0) {
			fnAppSync(2);
            Actions act = new Actions(wDriver);
            WebElement elt = wDriver.findElement(By.xpath(strObj));
            act.moveToElement(elt).build().perform(); 
            fnAppSync(2); }	}

	//Function to get the given object attribute value================================================================================================================================
	public String fnGetObjectAttributeValue(String strObject, String strAttribute) throws Throwable {
		String strTemp = "";
		if (fnObjectExists(strObject, 1) != 0) {
			int strIntVal = 0;
			try { strTemp = wDriver.findElement(By.xpath(strObject)).getAttribute(strAttribute);
				if (strTemp == null) strTemp = "";
				log.info("Object Attribute Value -- Object ID - " + strObject + ", Property name - " + strAttribute + ", value from object - " + strTemp); }
				catch(Exception exp) {
					log.info(exp.getMessage());
					strIntVal = 1; }
			if (strIntVal == 1) {
				try {
					Thread.sleep(100);
					strTemp = wDriver.findElement(By.xpath(strObject)).getAttribute(strAttribute);
					log.info("Object Attribute Value -- Object ID - " + strObject + ", Property name - " + strAttribute + ", value from object - " + strTemp); }
				catch(Exception exp) {
					log.error("[ERROR]    ****** Failed to read the value from the Object ID - " + strObject);
					log.info(exp.getMessage()); 
					testAssert.fail(exp.getMessage()); } } }
		return strTemp; }

	//function to verify the value exists or not in the drop down/ select item====================================================================================================
	public boolean fnSelectValue(String strObject, String strFieldName, String strDataSheet, String strDataPointer, String strColName) throws Throwable {
		boolean strFound = false;
		String strValue = fnGetExcelData(pageObjects.strDefaultPath, strDataSheet, strDataPointer, strColName).trim();
		if (strValue.trim().isEmpty()==false) {
		try{
			if (fnObjectExists(strObject, 1) != 0) {
				Select drpDwnObject = new Select(wDriver.findElement(By.xpath(strObject)));
				drpDwnObject.selectByValue(strValue);
				fnAppSync(1);
				log.info("Successfully selected the value : " + strValue + " - Object : " + strObject);
				Reporter.log("Selected the value : " + strValue + " - Field : " + strFieldName);
				strFound = true; } }
		catch(Exception exp){
			log.error("[ERROR]    ****** Failed to select the value from the Object ID - " + strObject);
			log.info(exp.getMessage()); 
			testAssert.fail(exp.getMessage());	} }
		return strFound; }
	
	//function to verify the value exists or not in the drop down/ select item====================================================================================================
	public boolean fnSelectByKeys(String strObject, String strFieldName, String strDataSheet, String strDataPointer, String strColName) throws Throwable {
		boolean strFound = false;
		String strValue = fnGetExcelData(pageObjects.strDefaultPath, strDataSheet, strDataPointer, strColName).trim();
		if (strValue.trim().isEmpty()==false) {
		try{
			if (fnObjectExists(strObject, 1) != 0) {
				fnMouseMove(strObject);
				fnObjectClick(strObject);
				fnAppSync(1);
				fnObjectClick(strObject);
				fnAppSync(1);

				char[] crLst = strValue.toCharArray();
				for (char ch : crLst) {
					fnRobotValueEnter(ch); }

				fnAppSync(1);
				log.info("Successfully selected the value : " + strValue + " - Object : " + strObject);
				Reporter.log("Selected the value : " + strValue + " - Field : " + strFieldName);
				strFound = true; } }
		catch(Exception exp){
			log.error("[ERROR]    ****** Failed to select the value from the Object ID - " + strObject);
			log.info(exp.getMessage()); 
			testAssert.fail(exp.getMessage());
			pageObjects.stepStatus = false; } }
		return strFound; }
	
	//function to verify the value exists or not in the drop down/ select item====================================================================================================
	public boolean fnObjectSendKeys(String strObject, String strFieldName, String strDataSheet, String strDataPointer, String strColName) throws Throwable {
		boolean strFound = false;
		String strValue = fnGetExcelData(pageObjects.strDefaultPath, strDataSheet, strDataPointer, strColName).trim();
		if (strValue.trim().isEmpty()==false) {
		try{
			if (fnObjectExists(strObject, 1) != 0) {
				wDriver.findElement(By.xpath(strObject)).sendKeys(strValue);
				fnAppSync(1);
				log.info("Successfully entered the value : " + strValue + " - Object : " + strObject);
				Reporter.log("Entered the value : " + strValue + " - Field : " + strFieldName);
				strFound = true; } }
		catch(Exception exp){
			log.error("[ERROR]    ****** Failed to enter the value from the Object ID - " + strObject);
			log.info(exp.getMessage()); 
			testAssert.fail(exp.getMessage());	
			pageObjects.stepStatus = false; } }
		return strFound; }
	
	//function to verify the value exists or not in the drop down/ select item====================================================================================================
	public boolean fnObjectCheck(String strObject, String strFieldName, String strDataSheet, String strDataPointer, String strColName) throws Throwable {
		boolean strFound = false;
		String strValue = fnGetExcelData(pageObjects.strDefaultPath, strDataSheet, strDataPointer, strColName).trim();
		if (strValue.trim().isEmpty()==false) {
		try{
			if (fnObjectExists(strObject, 1) != 0) {
				if(strValue.trim().toLowerCase().contentEquals("check")) {
					if(wDriver.findElement(By.xpath(strObject)).isSelected()==false) {
						wDriver.findElement(By.xpath(strObject)).click(); } }
				else if(strValue.trim().toLowerCase().contentEquals("uncheck")) {
					if(wDriver.findElement(By.xpath(strObject)).isSelected()==true) {
						wDriver.findElement(By.xpath(strObject)).click(); } }
				
				fnAppSync(1);
				log.info("Successfully checked the value : " + strValue + " - Object : " + strObject);
				Reporter.log("Entered the value : " + strValue + " - Field : " + strFieldName);
				strFound = true; } }
		catch(Exception exp){
			log.error("[ERROR]    ****** Failed to check the value from the Object ID - " + strObject);
			log.info(exp.getMessage()); 
			testAssert.fail(exp.getMessage());	
			pageObjects.stepStatus = false; } }
		return strFound; }

	// Function to enter the values in the input tag field ============================================================
	public void fnRobotValueEnter(char strValue) throws Throwable {

		Robot robot = new Robot();
		switch (Character.toUpperCase(strValue)) {
		case '0':
			robot.keyPress(KeyEvent.VK_0);
			robot.keyRelease(KeyEvent.VK_0);
			break;
		case '1':
			robot.keyPress(KeyEvent.VK_1);
			robot.keyRelease(KeyEvent.VK_1);
			break;
		case '2':
			robot.keyPress(KeyEvent.VK_2);
			robot.keyRelease(KeyEvent.VK_2);
			break;
		case '3':
			robot.keyPress(KeyEvent.VK_3);
			robot.keyRelease(KeyEvent.VK_3);
			break;
		case '4':
			robot.keyPress(KeyEvent.VK_4);
			robot.keyRelease(KeyEvent.VK_4);
			break;
		case '5':
			robot.keyPress(KeyEvent.VK_5);
			robot.keyRelease(KeyEvent.VK_5);
			break;
		case '6':
			robot.keyPress(KeyEvent.VK_6);
			robot.keyRelease(KeyEvent.VK_6);
			break;
		case '7':
			robot.keyPress(KeyEvent.VK_7);
			robot.keyRelease(KeyEvent.VK_7);
			break;
		case '8':
			robot.keyPress(KeyEvent.VK_8);
			robot.keyRelease(KeyEvent.VK_8);
			break;
		case '9':
			robot.keyPress(KeyEvent.VK_9);
			robot.keyRelease(KeyEvent.VK_9);
			break;
		case 'A':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_A);
			robot.keyRelease(KeyEvent.VK_A);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'B':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_B);
			robot.keyRelease(KeyEvent.VK_B);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'C':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_C);
			robot.keyRelease(KeyEvent.VK_C);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'D':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_D);
			robot.keyRelease(KeyEvent.VK_D);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'E':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_E);
			robot.keyRelease(KeyEvent.VK_E);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'F':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_F);
			robot.keyRelease(KeyEvent.VK_F);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'G':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_G);
			robot.keyRelease(KeyEvent.VK_G);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'H':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_H);
			robot.keyRelease(KeyEvent.VK_H);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'I':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_I);
			robot.keyRelease(KeyEvent.VK_I);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'J':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'K':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_K);
			robot.keyRelease(KeyEvent.VK_K);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'L':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_L);
			robot.keyRelease(KeyEvent.VK_L);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'M':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_M);
			robot.keyRelease(KeyEvent.VK_M);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'N':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_N);
			robot.keyRelease(KeyEvent.VK_N);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'O':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_O);
			robot.keyRelease(KeyEvent.VK_O);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'P':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_P);
			robot.keyRelease(KeyEvent.VK_P);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'Q':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_Q);
			robot.keyRelease(KeyEvent.VK_Q);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'R':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_R);
			robot.keyRelease(KeyEvent.VK_R);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;

		case 'S':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_S);
			robot.keyRelease(KeyEvent.VK_S);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'T':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_T);
			robot.keyRelease(KeyEvent.VK_T);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'U':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_U);
			robot.keyRelease(KeyEvent.VK_U);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'V':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_V);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'W':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_W);
			robot.keyRelease(KeyEvent.VK_W);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'X':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_X);
			robot.keyRelease(KeyEvent.VK_X);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'Y':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_Y);
			robot.keyRelease(KeyEvent.VK_Y);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case 'Z':
			robot.keyPress(KeyEvent.VK_SHIFT);
			robot.keyPress(KeyEvent.VK_Z);
			robot.keyRelease(KeyEvent.VK_Z);
			robot.keyRelease(KeyEvent.VK_SHIFT);
			break;
		case ' ':
			robot.keyPress(KeyEvent.VK_SPACE);
			robot.keyRelease(KeyEvent.VK_SPACE);
			break;
		case ',':
			robot.keyPress(KeyEvent.VK_COMMA);
			robot.keyRelease(KeyEvent.VK_COMMA);
			break;
		case '.':
			robot.keyPress(KeyEvent.VK_PERIOD);
			robot.keyRelease(KeyEvent.VK_PERIOD);
			break;
		case '_':
			robot.keyPress(KeyEvent.VK_UNDERSCORE);
			robot.keyRelease(KeyEvent.VK_UNDERSCORE);
			break; } }
	
	public void fnCaptureScreenShot(String stMsgValue) throws Throwable {
		File src= ((TakesScreenshot)wDriver).getScreenshotAs(OutputType.FILE);
		try {
			File screenShotName = new File("C:/ethoca_screens/screen_" + pageObjects.intImgCnt + ".jpg");
			FileUtils.copyFile(src, screenShotName);
			pageObjects.intImgCnt = pageObjects.intImgCnt + 1; 

			String filePath = screenShotName.toString();
			strScreenPath = filePath;
			pageObjects.strImgList[0][intImgCnt] = filePath;
			pageObjects.strStepList[0][intImgCnt] = stMsgValue;
			if (stepStatus == true)
				pageObjects.strStepListStatus[0][intImgCnt] = "true";
			else
				pageObjects.strStepListStatus[0][intImgCnt] = "false";

			String path = "<img src=\"file://" + filePath + "\" alt=\"\"/>";
			Reporter.log(path);
			fnScreenCapture(wDriver, path);	}
		catch (IOException e) {
		  log.error(e.getMessage());
		  Reporter.log(e.getMessage()); } }

	//function to capture the screen============================================================================================================================================
	public void fnScreenCapture1(String strFileName) throws Throwable {
		File scrFile;
		String fileName="";
		strFileName = pageObjects.strScreenPath;
		
		pageObjects.strImgList[0][pageObjects.intImgCnt] = fileName;
        pageObjects.strStepList[0][pageObjects.intImgCnt] = pageObjects.strGVStepName;

        if ((pageObjects.strScnStatus == true) || (pageObjects.strGVStepName.trim().toLowerCase().contains("success")))
        	pageObjects.strStepListStatus[0][pageObjects.intImgCnt] = "true";
        else
        	pageObjects.strStepListStatus[0][pageObjects.intImgCnt] = "false";

        pageObjects.intImgCnt++;
        pageObjects.strGVStepName = ""; }
	
	//function to generate the HTML report for test execution =================================================================================================================
	public synchronized void fnGenerateHTMLHeaderReport(String strHTMLFilePath) throws Throwable, Throwable {
		//Steps to create the HTML report header section
		FileWriter strHTMLFile;
		String strFoldPath = "\\" + (pageObjects.strHTMLPath).replace("\\\\", "\\");

		String strCurrTime1 = new SimpleDateFormat("HH.mm").format(new Date());
		String strHeader = "<html><head> <style> a:link { color: black; } a:visited { color: black; } a:hover { color: Blue; } a:active { color: blue; } </style> </head>"
				+ "<font face='verdana'><h4 align='center'> ETHOCA TEST - Automation Execution Report" + "</h4><hr><td>"
						+ "<h6 align='left'><font color='blue' SIZE=2> URL : " + pageObjects.strEnvironmentURL + "</font> </h6></td>"
								
										+ "<h6 align ='left'><font color='blue' SIZE=2>"
										+ "<a href=" + "\"" + strFoldPath + "\"" +  "onclick = open(this.href, this.target, 'width=600, height=450, top=200, left=250'); return false; target=_blank> Click to open Results Folder.. ..</a>"
										+ "</font></h6> "
										+ "<h6><font color='black' SIZE=2> Product : " + pageObjects.strProductName + " </h6> "
										
										+ "<TABLE BORDER=2 BGCOLOR=#ffff00>"
										+ "<TR><TH><FONT COLOR=BLACK FACE='Geneva, Arial' SIZE=2>Use_Case/Story_Name_________</FONT></TH>"
										+ "<TH><FONT COLOR=BLACK FACE='Geneva, Arial' SIZE=2>Status</TH>"
										+ "<TH><FONT COLOR=BLACK FACE='Geneva, Arial' SIZE=2>TimeStamp</TH>"
										+ "<TH><FONT COLOR=BLACK FACE='Geneva, Arial' SIZE=2>Remarks______________________________________________</TH></TR>";
		pageObjects.strHTMLName = "Ethoca_AutRep_Sample_Report.HTML";

		File strFile = new File(pageObjects.strHTMLPath + pageObjects.strHTMLName);
		if (strFile.exists() == true) {
			strFile.delete();
			Thread.sleep(100);	}
		
		if (strFile.exists() == false) {
			strFile.createNewFile();
			strHTMLFile = new FileWriter(pageObjects.strHTMLPath + pageObjects.strHTMLName, true);
			strHTMLFile.write(strHeader);
			strHTMLFile.close(); }	}

	//function to generate the HTML results row one by one ===================================================================================
	public synchronized void fnGenerateHTMLRows(String strHTMLFilePath, String strTC, String strStatus, String strDealID, String strTime, String strRemarks) throws Throwable{
		String strColor=null;
		String strTempPath = pageObjects.strScreenPath;
		String strTempPath1 = "\\\\";
		Character chr1 = (char) 92;
		int intTemp=0;
		for (Character chr : strTempPath.toCharArray()){
			if (intTemp == 0) {
				chr1 = chr;
				intTemp = 1; }
			else {
				if (((int) chr == 92) && ((int)chr1 == 92)){}
				else if (((int) chr == 92) && ((int)chr1 != 92))
					strTempPath1 = strTempPath1 + "\\";
				else
					strTempPath1 = strTempPath1 + chr; }
			chr1=chr; }

		pageObjects.strPDFPath =  strTC + ".PDF";
	
		if (strStatus.trim().toLowerCase().contentEquals("pass") == true)
			strColor = "#006633"; // green color code
		else if (strStatus.trim().toLowerCase().contentEquals("fail") == true)
			strColor = "#D80000"; //red color code
		else if(strStatus.trim().toLowerCase().contentEquals("pass&fail") == true)
			strColor = "#FF69B4"; // hotpink color code

		String strCurrTime1 = new SimpleDateFormat("HH.mm").format(new Date());
		String strRows = "<TR  style='background:#f2f2f2'" + ";text-align='left'>";

		if(pageObjects.stepStatus == true){
			strRows = strRows + "<TH BGCOLOR = '' ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=1> <a href=" + "\"" + pageObjects.strPDFPath + "\"" + "onclick = open(this.href, this.target, 'width=600, height=450, top=200, left=250'); return false; target=_blank><font color='" + strColor + "'>" + strTC + "</font></a> </FONT></TH>";
		}else{ 
			strRows = strRows + "<TH BGCOLOR = '' ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=1> <a href=" + "\"" + pageObjects.strPDFPath + "\"" + "onclick = open(this.href, this.target, 'width=600, height=450, top=200, left=250'); return false; target=_blank><font color='" + strColor + "'>" + strTC + "</font></a> </FONT></TH>"; }
		strRows = strRows + "<TH BGCOLOR = '' ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=1><font color='" + strColor + "'>" + strStatus + "</font> </FONT></TH>";

		strRows = strRows + "<TH ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=1>" + strCurrTime1 + " </FONT></TH>";
		strRows = strRows + "<TH STYLE='WORD-WRAP:BREAK-WORD' ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=1>" + strRemarks + " </FONT></TH></TR>";

		File strFile = new File(pageObjects.strHTMLPath + pageObjects.strHTMLName);
		if (strFile.exists() == true) {
			FileWriter strHTMLFile = new FileWriter(pageObjects.strHTMLPath + pageObjects.strHTMLName, true);
			strHTMLFile.write(strRows);
			strHTMLFile.close();  } }
	
	
	//Function to prepare the PDF doc with the test case screen shots==============================================================================================================
	public void fnPDFScreenReport(String dataPointer) throws Throwable {
		if (pageObjects.strGenPDF == true) {
			Document document = new Document(PageSize.A4);
			Font font;
			PdfWriter pdfWrite = PdfWriter.getInstance(document, new FileOutputStream(strHTMLPath + dataPointer.trim() + ".pdf"));
			document.open();
			PdfContentByte cb = pdfWrite.getDirectContent();
			ColumnText ct = new ColumnText(cb);

			int intCnt=0;
			for (intCnt = 0; intCnt < pageObjects.strImgList[0].length; intCnt++){
				String imgFileName = pageObjects.strImgList[0][intCnt];
				try{
					if (imgFileName != null && !imgFileName.isEmpty()) {
							Image imag1 = Image.getInstance(imgFileName);
							document.setPageSize(imag1);
							document.newPage();
							if (pageObjects.strStepListStatus[0][intCnt].contains("true"))
								font = new Font(FontFamily.COURIER,25.0f,Font.BOLD,BaseColor.GREEN);
							else
								font = new Font(FontFamily.COURIER,25.0f,Font.BOLD,BaseColor.RED);
								font.setStyle(Font.UNDERLINE);
								Paragraph para = new Paragraph(pageObjects.strStepList[0][intCnt], font);
								para.setAlignment(Paragraph.ALIGN_TOP);
								para.setAlignment(Paragraph.ALIGN_CENTER);
								imag1.setAbsolutePosition(0, 0);
								Phrase ph = new Phrase("test");
								ct.setText(ph);
								document.add(para);
								document.add(imag1); } }
				catch (Exception exp){
					System.out.println("[ERROR]    Unable to capture/ attach the PDF to report : " + imgFileName);	} }

			String imgFileName="";
			for (intCnt = 0; intCnt < pageObjects.strImgList[0].length; intCnt++){
				imgFileName = pageObjects.strImgList[0][intCnt];
	
				//Step to delete the image file on exists.
				if (imgFileName != null)
					 if (imgFileName.isEmpty() == false) { 
						 File srcFile = new File(imgFileName);
						 if (srcFile.exists() == true)
							 srcFile.delete(); } }
			pageObjects.strGenPDF = false;

			document.close(); }  }
	
	
	
	
 }