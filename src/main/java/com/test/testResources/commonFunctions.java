/*

package automation_resources;
import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.event.KeyEvent;
import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintStream;
import java.math.BigDecimal;
import java.net.InetAddress;
import java.nio.channels.FileChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.List;
import java.util.Properties;
import java.util.Scanner;
import java.util.concurrent.TimeUnit;
import javax.mail.Message;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.parser.JSONParser;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.Font.FontFamily;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfWriter;
import com.thoughtworks.selenium.Selenium;

import external_actions.irs_dealBooking;
import external_actions.loginLogout;
import junit.framework.Assert;
import jxl.write.DateTime;

import java.awt.datatransfer.StringSelection;

 @SuppressWarnings("unused")
public class commonFunctions extends pageObjects {

	//function to capture the screen============================================================================================================================================
	public void fnScreenCapture(WebDriver wDriver, String strFileName) throws Throwable {
		File scrFile;
		String fileName="";
		
		if (pageObjects.strScrnFileName == false) {
			strFileName = pageObjects.strScreenFileName;
			pageObjects.strScreenFileCnt = pageObjects.strScreenFileCnt + 1;
			strFileName = pageObjects.strScreenPath1 + strFileName; }
		
		if (pageObjects.strScreenCapture == true) {
		    //Robot robot = new Robot();
			System.out.println("[INFO]   taking screen capture in fnScreenCapture - " + strFileName);
	        String format = ".jpg";
	        if (pageObjects.strScrnFileName == false)
	        	fileName = strFileName + pageObjects.intImgCnt + format;
	        else
	        	fileName = strFileName + format;
			try{ File file = new File(fileName);
				File myFile = new File(file, fileName);
				if (myFile.exists() == true)
					myFile.delete(); }  catch(Exception e) { System.out.println("[ERROR]    Unable to delete the file ... " + fileName);}
			try{
				scrFile = ((TakesScreenshot)wDriver).getScreenshotAs(OutputType.FILE);
				FileUtils.copyFile(scrFile, new File(fileName)); } 
			catch (Exception exp) {
				try {
				System.out.println("[INFO]   taking screen capture in catch statement - fnScreenCapture - " + strFileName);
				scrFile = ((TakesScreenshot)wDriver).getScreenshotAs(OutputType.FILE); }
				catch (Exception exp1){
					pageObjects.strScnStatus = false;
					System.out.println("[INFO]   Failed to capture the screen, no browser exists - fnScreenCapture - " + strFileName); } }

	        pageObjects.strImgList[0][pageObjects.intImgCnt] = fileName;
	        pageObjects.strStepList[0][pageObjects.intImgCnt] = pageObjects.strGVStepName;
	
	        if ((pageObjects.strScnStatus == true) || (pageObjects.strGVStepName.trim().toLowerCase().contains("success")))
	        	pageObjects.strStepListStatus[0][pageObjects.intImgCnt] = "true";
	        else
	        	pageObjects.strStepListStatus[0][pageObjects.intImgCnt] = "false";
	
	        pageObjects.intImgCnt++;
	        pageObjects.strGVStepName = ""; } }

	//function to wait/ sync for given amount of time in seconds==================================================================s==============================================
	public void fnAppSync(int intTime) throws Exception {
		Robot robot = new Robot();
		//TimeUnit.SECONDS.sleep(intTime);
		robot.delay(intTime * 500); }
	
	//Function to click on the element================================================================================================================================
	public void fnObjectClick(String strObj) throws Throwable{
		fnAppSync(pageObjects.gnWaitTime);
		pageObjects.strPrntMsg = false;
		if (fnObjectExists(strObj, 1) != 0) {
			try{ wDriver.findElement(By.xpath(strObj)).click();	
				System.out.println("[INFO]    1. Object Click successful - " + strObj);} 
			catch (Exception exp) {
				System.out.println("[INFO]    2. Trying to click on - " + strObj);
				fnAppSync(50);
				try {
					wDriver.findElement(By.xpath(strObj)).click(); }
				catch (Exception exp1) {
					pageObjects.strScnStatus = false;
					System.out.println("[ERROR]   UnExpected Error - Unable to click on the object location -> " + strObj);
					pageObjects.strGVStepName = "UnExpected Error - Unable to click on the object location -> " + strObj;
					pageObjects.strGVErrMsg = "UnExpected Error - Failed to Find/Click on the link/button";
					fnScreenCapture(wDriver, strScreenPath.concat("dpmDealConfirmFailBkRnDlFailure"));	} } } }
	
	public void fnObjectHighLight(String strObj) throws Throwable {
		WebElement ele = wDriver.findElement(By.xpath(strObj));		        
		JavascriptExecutor js = (JavascriptExecutor) wDriver;
		js.executeScript("arguments[0].setAttribute('style', 'background: ; border: 1px solid red;');", ele);
		Thread.sleep(100);
		js.executeScript("arguments[0].setAttribute('style', 'background: ; border: ;');", ele);
	}
	
	//Function to check and click the element================================================================================================================================
	public void fnObjectCheckAndClick(String strObj, int intAttempts) throws Throwable {
		fnAppSync(pageObjects.gnWaitTime);
		pageObjects.strPrntMsg = false;
		int objFound = 0;
		Robot robot = new Robot();
		for (int intTemp=0;intTemp<intAttempts;intTemp++) {
			if (fnObjectExists(strObj, 1) != 0) {
				try{ wDriver.findElement(By.xpath(strObj)).click();	
					System.out.println("[INFO]   Object Check & Click successful1 - " + strObj);
					objFound = 1;} 
				catch (Exception exp) {
					System.out.println("[ERROR]   Try1 : " + intTemp + " - Unable to click on the object location -> " + strObj);
					robot.delay(3000);
					try {
						wDriver.findElement(By.xpath(strObj)).click();
						System.out.println("[INFO]   Object Check&Click successful2 - " + strObj);
						objFound = 1; }
					catch (Exception exp1) {
						System.out.println("[ERROR]   Try2 : " + intTemp + " - Unable to click on the object location -> " + strObj); } } }

		if (objFound == 1) 
			intTemp=intAttempts+5;
		else if (objFound == 0)
			robot.delay(8000); }
		
		if (objFound == 0) {
			pageObjects.strScnStatus = false;
			System.out.println("[ERROR]   UnExpected Error - Unable to click on the object location -> " + strObj);
			pageObjects.strGVStepName = "UnExpected Error - Unable to click on the object location -> " + strObj;
			pageObjects.strGVErrMsg = "UnExpected Error - Failed to Find/Click on the link/button";
			fnScreenCapture(wDriver, strScreenPath.concat("dpmDealObjectNotFound")); } }

	
	//Function to double click on the element================================================================================================================================
	public void fnObjectDblClick(String strObj) throws Throwable{
		fnAppSync(pageObjects.gnWaitTime);
		if ((wDriver.findElements(By.xpath(strObj)).size()) ==  1) {
			try {
				WebElement wButton = wDriver.findElement(By.xpath(strObj));
		        Actions action = new Actions (wDriver);
		        action.moveToElement(wButton).doubleClick().build().perform(); }
			catch(Exception exp) {
				System.out.println("[ERROR]    Unable to double click on object - " + strObj);
			}
		} else
			System.out.println("[ERROR]    ****** Failed to find the object with ID " + strObj); }
	
	//Function to select the check box to checked or unchecked================================================================================================================================
	public void fnObjectSelectCheckBox(String strObj, String strCheck) throws Throwable{
		if (fnObjectExists(strObj, 1) != 0) {
			boolean strChkStat = wDriver.findElement(By.xpath(strObj)).isSelected();
			if (strCheck.trim().toLowerCase().contentEquals("check") == true)
				if (strChkStat == false)
					fnObjectClick(strObj);
			else if (strCheck.trim().toLowerCase().contentEquals("uncheck")  == true)
				if (strChkStat == true)
						fnObjectClick(strObj); } }
	
	//Function to select the check box to checked or unchecked================================================================================================================================
	public void fnObjectSendKeys(String strObj, String strVal) throws Throwable{
		if (fnObjectExists(strObj, 1) != 0) {
			String strChkStat = fnGetObjectAttributeValue(strObj, "readonly");
			if ((strChkStat.toLowerCase().contentEquals("null")) || (strChkStat.toLowerCase().contentEquals("false") == true)) {
					wDriver.findElement(By.xpath(strObj)).click();
					wDriver.findElement(By.xpath(strObj)).clear();
					fnAppSync(1);
					wDriver.findElement(By.xpath(strObj)).sendKeys(strVal);
					wDriver.findElement(By.xpath(strObj)).sendKeys(Keys.TAB); } } }
	
	//Function to get the given object exists or not================================================================================================================================
	public int fnObjectExists(String strObject, int timeInterval) throws Throwable {
		Robot robot = new Robot();
		int strDisp = 0;
		for(int intVal=0; intVal <= timeInterval; intVal++) {
			try {
				if (wDriver.findElements(By.xpath(strObject)).size() > 0)
					if (wDriver.findElement(By.xpath(strObject)).isDisplayed() == true) {
								strDisp = 1; } } 
			catch (Exception exp) { strDisp = 0; }
			if (strDisp == 1) break;
			pageObjects.totWaitTime = pageObjects.totWaitTime + 1;
			//fnAppSync(pageObjects.gnWaitTime);
			robot.delay(100); }
		
			if (pageObjects.strPrntMsg == true)
				System.out.println("[INFO]   Object ID - " + strObject + " No. of Objects Found : " + strDisp);
			else pageObjects.strPrntMsg = true;

		if (strDisp == 1) {
			fnObjectHighLight(strObject); }
		return strDisp; }
	
	//Function move the mouse on to object and click =================================================================================================================================
	public void fnMouseMoveAndClick(String strObj) throws Throwable {
		fnAppSync(pageObjects.gnWaitTime);
		pageObjects.strPrntMsg = false;
		if (fnObjectExists(strObj, 1) != 0) {
			try{ wDriver.findElement(By.xpath(strObj)).click();	
				System.out.println("[INFO]    1. Object Click successful - " + strObj);} 
			catch (Exception exp) {
				System.out.println("[INFO]    2. Trying to click on - " + strObj);
				fnAppSync(50);
				try {
					WebElement el = wDriver.findElement(By.xpath(strObj));
				    Actions builder = new Actions(wDriver);
				    builder.moveToElement(el).click(el);
				    builder.perform(); }
				catch (Exception exp1) {
					pageObjects.strScnStatus = false;
					System.out.println("[ERROR]   UnExpected Error - Unable to move and click on the object location -> " + strObj);
					pageObjects.strGVStepName = "UnExpected Error - Unable to move and click on the object location -> " + strObj;
					pageObjects.strGVErrMsg = "UnExpected Error - Failed to Find/Click on the link/button";
					fnScreenCapture(wDriver, strScreenPath.concat("dpmDealConfirmFailBkRnDlFailure"));	} } } }
	
	//Function to get the given object attribute value================================================================================================================================
	public String fnGetObjectAttributeValue(String strObject, String strAttribute) throws Throwable {
		String strTemp = "";
		if (fnObjectExists(strObject, 1) != 0) {
			int strIntVal = 0;
			try { strTemp = wDriver.findElement(By.xpath(strObject)).getAttribute(strAttribute);
				if (strTemp == null) strTemp = "";
				System.out.println("[INFO]   Object Attribute Value -- Object ID - " + strObject + ", Property name - " + strAttribute + ", value from object - " + strTemp); }
				catch(Exception exp) { strIntVal = 1; }
			if (strIntVal == 1) {
				try {
					fnAppSync(5);
					strTemp = wDriver.findElement(By.xpath(strObject)).getAttribute(strAttribute);
					System.out.println("[INFO]   Object Attribute Value -- Object ID - " + strObject + ", Property name - " + strAttribute + ", value from object - " + strTemp); }
				catch(Exception exp) {
					System.out.println("[ERROR]    ****** Failed to read the value from the Object ID - " + strObject); } } }
			//if (strTemp == null) strTemp="";
		return strTemp; }

	//Function get the values from excel sheet and compare against the application================================================================================================================================
	public void fnDealCompareValues(String strDataSheetPath,String dataSheet,String dataPointer, String strFieldName, String strColname, String strObject,String strAttribute) throws Throwable {
		String strTempVal = "";
		//step to retrieve the trade date from excel and book detail screen and compare
		String strVal = fnGetExcelData(strDataSheetPath, dataSheet, dataPointer, strColname);
		String strTemp = strVal.trim();

		if (strVal.isEmpty() != true) {
			if (strVal.trim().toLowerCase().contentEquals("today") == true)
				strVal = fnGetCurrentDate("yyyy-mm-dd");
			
			if ((isNumeric(strVal) == true) && (strVal.contains(".") == false))
				strVal = strVal.replaceAll(",", "") + ".00";
			else
				strVal = (strVal.toLowerCase()).replaceAll(",", "");
			
			//Step to get the text from the SPAN object type
			String strNodeVal = fnGetObjectAttributeValue(strObject, "nodeName");
			if (strNodeVal.trim().length()>0) {
				if (strNodeVal.trim().toLowerCase().contentEquals("span") == true)
					strAttribute = "outerText";
				
				//step to read the object attribute property value
				strTempVal = fnGetObjectAttributeValue(strObject, strAttribute);
				
				if (strNodeVal.trim().toLowerCase().contentEquals("select") == true) {
					if (isNumeric(strTempVal) == true)
						strTempVal = fnGetListElementText(strObject, Integer.valueOf(strTempVal)); }
				
				String strTemp1 = strTempVal;
				if ((isNumeric(strTempVal) == true) && (strTempVal.contains(".") == false))
					strTempVal = strTempVal.replaceAll(",", "") + ".00";
				else
					strTempVal = ((strTempVal.toLowerCase()).replaceAll(",", "")).trim();

				if ((strVal.toLowerCase().trim().contentEquals(strTempVal.trim().toLowerCase())) == true)
					System.out.println("[INFO]    == == == Comparison passed - " + strFieldName);
				else {
					System.out.println("[ERROR]    ****** Comparison Failed - " + strFieldName);
					pageObjects.strCompareVal = false;
					if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
						pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + "Field name : " + strFieldName + " ,  Expected : " + strTemp + " ,  Actual : " + strTemp1 + "; \n ";
				} } 
			else {
				System.out.println("[ERROR]    ****** Failed to retrieve the value from field : " + strFieldName);
				if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
				pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + "Failed to retrieve the value from field : " + strFieldName + "; \n "; } }
		}

	
	//Function to book a deal after "SaveAs" ================================================================================================================================
		public void dpmSaveAsDealBook(String dataSheet, String dataPointer) throws Throwable {

			String strTxt = new String();
			//Click on Book deal button		
			if(pageObjects.strFunctionality.trim().toLowerCase().contains("loan") || pageObjects.strFunctionality.trim().toLowerCase().contains("deposit")){
				if (fnObjectExists(objBtnBook, 2) != 0)
					fnObjectClick(objBtnBook);

				if (fnObjectExists(objBtnBook, 2) != 0)
					fnObjectClick(objBtnBook); }
			else {
				if (fnObjectExists(objBtnBook, 2) != 0)
					fnObjectClick(objBtnBook); }

			//Click on confirmation button
			fnAppSync(1);
			if (fnObjectExists(objConfirmAction, 2) != 0) {
				fnObjectClick(objConfirmAction);
				fnAppSync(1); }

			int strFound = -1;
			boolean strSTxt = false;
			for(int intTemp = 0; intTemp<20; intTemp++) {
				strFound = -1;
				if ((fnObjectExists(objIDSuccessText, 1) != 0) || (fnObjectExists(objIDInBookingText, 1) != 0))
				{ 	String strTxt1="";
					if (fnObjectExists(objIDSuccessText, 1) != 0){
						Point objCoordinates = wDriver.findElement(By.xpath(objIDSuccessText)).getLocation();
						Robot robot = new Robot();
						robot.mouseMove(objCoordinates.getX()+66, objCoordinates.getY()+86);
						strTxt = wDriver.findElement(By.xpath(objIDContent)).getText();
						strSTxt = (wDriver.findElement(By.xpath(objIDSuccessText)).getText().toLowerCase().contains("success"));
						robot.mouseMove(objCoordinates.getX()+65, objCoordinates.getY()+84);
						strTxt1 = wDriver.findElement(By.xpath(objIDContent)).getText(); }
					else if (fnObjectExists(objIDInBookingText, 1) != 0){
						Point objCoordinates = wDriver.findElement(By.xpath(objIDInBookingText)).getLocation();
						Robot robot = new Robot();
						robot.mouseMove(objCoordinates.getX()+66, objCoordinates.getY()+86);
						strTxt = wDriver.findElement(By.xpath(objIDInBookingContent)).getText();
						strSTxt = (wDriver.findElement(By.xpath(objIDInBookingText)).getText().toLowerCase().contains("success"));
						robot.mouseMove(objCoordinates.getX()+65, objCoordinates.getY()+84);
						strTxt1 = wDriver.findElement(By.xpath(objIDInBookingContent)).getText(); }
					if (strTxt1.trim().toLowerCase().contains("object") == true) {
						strTxt = "[Object Object]";
						strFound = 0; }
					else 
						strFound = 1; }
				else if (fnObjectExists(objIDErrText, 1) != 0)
				{	Point objCoordinates = wDriver.findElement(By.xpath(objIDErrText)).getLocation();
					Robot robot = new Robot();
					robot.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+87);
					strTxt = wDriver.findElement(By.xpath(objServerErrorContent)).getText();
					strSTxt = (wDriver.findElement(By.xpath(objServerErrorContent)).getText().toLowerCase().contains("success"));
					robot.mouseMove(objCoordinates.getX()+64, objCoordinates.getY()+83);
					strFound = 0; pageObjects.strScnStatus = false; }
					else if (fnObjectExists(objFormErr, 1) != 0)
					{	strFound = 0;
					strTxt = fnGetObjectAttributeValue(objFormErr, "outerText");
					pageObjects.strScnStatus = false;}
					fnAppSync(1);
					if (strFound != -1) break;}

			if (strFound == 1) {
				System.out.println("[INFO]   Deal Popup Message details - " + strTxt);
				if (strSTxt == true) {
					if (pageObjects.strDealAmend == false) {
						pageObjects.strGVStepName = "New Deal Booked Successfull and Deal Id - " + strTxt;
						System.out.println("[INFO]   New Deal Booked Successfull and Deal Id - " + strTxt);
					} else if (pageObjects.strDealAmend == true) {
						pageObjects.strGVStepName = "Deal Amendment Successfull and Deal Id - " + strTxt;
						System.out.println("[INFO]   Deal Amendment Successfull and Deal Id - " + strTxt); }
					pageObjects.strGVDealID = strTxt;
					fnScreenCapture(wDriver, strScreenPath.concat("dpmDealBookSuccess") + "-"+ dataPointer);
					
					fnObjectClick(objPXVLabel);
					if (fnObjectExists(objIDSuccessText, 2)!=1) {
						fnObjectClick(objIDSuccessText);
						fnObjectClick(objIDSuccessText); }
					
					//step to update the deal id into the excel sheet
					fnUpdateExcelCellData(strDataSheetPath, dataSheet, dataPointer, "DPM_Deal_ID", strTxt);
					String strTime = fnGetCurrentDate("yyyy/mm/dd") + " " + fnGetCurrentTime();
					fnUpdateExcelCellData(strDataSheetPath, dataSheet, dataPointer, "Deal_DateTime", strTime); }
				else
					strFound = 0; }

			if (strFound == 0) {	
				pageObjects.strScnStatus = false;
				if (pageObjects.strDealAmend == false) {
					pageObjects.strGVErrMsg = "Failed to Book the " + pageObjects.strProductName + " deal, error msg - " + strTxt;
					pageObjects.strGVStepName = "Failed to Book the " + pageObjects.strProductName + " deal, error msg - " + strTxt;
					fnScreenCapture(wDriver, strScreenPath.concat("dpmDealBookFailure") + "-"+ dataPointer);
					System.out.println("[ERROR]   Failed to Book the " + pageObjects.strProductName + " deal, error msg - " + strTxt); }
				else if (pageObjects.strDealAmend == true) {
					pageObjects.strGVErrMsg = "Failed to Amend the " + pageObjects.strProductName + " deal, error msg - " + strTxt;
					pageObjects.strGVStepName = "Failed to Amend the " + pageObjects.strProductName + " deal, error msg - " + strTxt;
					fnScreenCapture(wDriver, strScreenPath.concat("dpmDealBookFailure") + "-"+ dataPointer);
					System.out.println("[ERROR]   Failed to Amend the " + pageObjects.strProductName + " deal, error msg - " + strTxt); } } 
		}

	//Function to get the data from the excel file================================================================================================================================
	@SuppressWarnings("deprecation")
	public synchronized String fnGetExcelData(String fileName, String sheetName, String dataPointer, String columnName) throws Throwable {
		String cellValue = "";
		//System.out.println("fnGetExcel Data --- " + fileName);
		//get the rows count from the excel sheet
		int rowNum = fnGetExcelRowNum(fileName, sheetName, dataPointer);		
		if (rowNum == -1){
			System.out.println("[WARNING]    Incorrect DataPointer - " + dataPointer); }
		
		//get the columns count from the excel sheet
		int colNum = fnGetExcelColNum(fileName, sheetName, columnName);
		if (colNum == -1){
			System.out.println("[WARNING]    Incorrect column name - " + columnName); }

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
	        		        	
	        	System.out.println("[INFO]   == Cell Row : " + rowNum + " Datapointer : " + dataPointer + " - Cell Column : " + colNum + ", Column Name : " + columnName + " - Cell Value : " + cellValue); }
	       else if ((rowNum == -1) || (colNum == -1)) {
	    	   cellValue = "";
	    	   //System.out.println("[INFO]    Incorrect excel column/ row values == Cell Row : " + rowNum + " - Cell Column : " + colNum + ", Column Name : " + columnName + " - Cell Value : " + cellValue); 
	    	   }
		return cellValue.trim(); }

	//method to get the data from the excel file using the row number and column name================================================================================================================================
	@SuppressWarnings({ "deprecation", "resource" })
	public synchronized String fnGetExcelCellData(String fileName, String sheetName, int strRowNum, String columnName) throws Throwable {
		String cellValue = "";
		//get the columns count from the excel sheet
		int colNum = fnGetExcelColNum(fileName, sheetName, columnName);
		if (colNum > -1 && strRowNum > -1) {
			File file=new File(fileName);
	        if(file.exists()){
		        InputStream myxls = new FileInputStream(new File(fileName));
		        HSSFWorkbook book = new HSSFWorkbook(myxls); 
		        HSSFSheet sheet = book.getSheet(sheetName);
		        Row row = sheet.getRow(strRowNum);
		        try{ Cell cell = row.getCell(colNum);	        
		        if (cell == null)
		        	cellValue = "";
		        else if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
		        	cellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
		        else if(cell.getCellType() == Cell.CELL_TYPE_STRING)
		        	cellValue = cell.getStringCellValue();
		        else if(cell.getCellType() == Cell.CELL_TYPE_BLANK)
		        	cellValue = ""; }
		        catch (Exception exp) {
		        	System.out.println("[ERROR]    Unable to read the value from ColumnName : " + columnName + ", Row. " + strRowNum + ", Sheet Name : " + sheetName + ", DataFile : " + fileName);
		        	cellValue = "";}
		        myxls.close(); } }
		else cellValue = "";
		return cellValue; }

	//function to get the rows count in the given data sheet================================================================================================================================
	@SuppressWarnings("resource")
	public synchronized static int fnGetExcelRowCount(String fileName, String sheetName) throws Exception {
		System.out.println("[INFO]   == Data File Name : " + fileName);
		System.out.println("[INFO]   == Data Sheet Name : " + sheetName);
		
		File file=new File(fileName);
		int rowCount = 0;
		try{
	        if(file.exists()){
		        InputStream myxls = new FileInputStream(fileName); 
		        HSSFWorkbook book = new HSSFWorkbook(myxls); 
		        HSSFSheet sheet = book.getSheet(sheetName);
		        rowCount = sheet.getLastRowNum();
		        myxls.close(); } }
		catch (Exception exp){
			System.out.println("[ERROR]    File not found .... " + fileName); }
		return rowCount; }

	//function to get the COLUMN count in the given data sheet=========================================================================================
	@SuppressWarnings("resource")
	public synchronized static int fnGetExcelColumnCount(String fileName, String sheetName) throws Exception {
		File file=new File(fileName);
		int colCount = 0;
        if(file.exists()){
	        InputStream myxls = new FileInputStream(fileName); 
	        HSSFWorkbook book = new HSSFWorkbook(myxls); 
	        HSSFSheet sheet = book.getSheet(sheetName);
	        colCount = sheet.getRow(pageObjects.intSheetRowNum).getPhysicalNumberOfCells(); 
	        myxls.close(); }
		return colCount; }
	
	//function to get the row number of the given data pointer====================================================================================================
	@SuppressWarnings({ "deprecation", "resource" })
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
	        	
	        	//System.out.println("Cell Value - " + strCellValue);
	        	if (strCellValue.trim().toLowerCase().contentEquals(dataPointer.trim().toLowerCase()))
	        		colNum = intLoop;
	        	if (colNum != -1) break; } }
		return colNum;	}
	
	//function to get the column number of the given datapointer====================================================================================================
	@SuppressWarnings({ "resource", "deprecation" })
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
	        myxls.close(); }    
		return colNum; }

	//function to close the opened widgets=============================================================================================================================
	public void fnCloseDPMWidgets(){
		try{
			if (wDriver.findElement(By.xpath(objTrdActvityClose)).isDisplayed() == true)
				wDriver.findElement(By.xpath(objTrdActvityClose)).click(); }
		catch(Exception e){} }

	// function to update excel cell data
	public synchronized void fnUpdateExcelCellData_BACKUP(String fileName, String sheetName, String dataPointer, String colName, String value) throws Throwable{
			//fnAppSync(5);
			try{
				System.out.println("File Name : " + fileName);
				System.out.println("Sheet Name : " + sheetName);
				System.out.println("Datapointer : " + dataPointer);
				System.out.println("Col. Name : " + colName);
				System.out.println("Col. value : " + value);
				
				File file = new File(fileName);
				if(file.exists()){
					HSSFWorkbook workBook = null;
					HSSFSheet sheet = null;
					HSSFRow row = null;
					HSSFCell cell = null;
					int rowNum = fnGetExcelRowNum(fileName, sheetName, dataPointer);
					int colNum = fnGetExcelColNum(fileName, sheetName, colName);

					if(rowNum > -1 && colNum >-1){
						FileInputStream fis = new FileInputStream(file);

						workBook = new HSSFWorkbook(fis);
						sheet = workBook.getSheet(sheetName.trim());
						if(sheet == null){
							sheet = workBook.createSheet(sheetName); }
						row = sheet.getRow(rowNum);
						if(row == null){
							row = sheet.createRow(rowNum);	}
						cell = sheet.getRow(rowNum).createCell(colNum);
						cell.setCellValue(value);
						fis.close();	}
					FileOutputStream outFile =new FileOutputStream(new File(fileName));
					workBook.write(outFile);
					workBook.close();
					outFile.close();
					//fnAppSync(1);
					boolean strFileLock=false;
					while (strFileLock == false){
						File file1 = new File(fileName);
						strFileLock = file1.renameTo(file);
						//fnAppSync(1); 
						}
					System.out.println("[INFO]    "+colName+" Data Updated Successfully in "+fileName);	}	}
			catch(Exception e){
				System.out.println("[ERROR] Exception Occured While updating "+colName+" Data in "+fileName);	}	}

//function to update the excel sheet cell data=============================================================================================================================
	@SuppressWarnings({ "deprecation", "resource" })
	public synchronized int fnUpdateExcelCellData_ORG(String fileName, String sheetName, String dataPointer, String columnName, String updValue) throws Throwable {
		int intreturnvalue = 0;
		try {
			String strOldFileName = fileName;
			if (fileName.trim().toLowerCase().contains(".xls")) {
				fileName = fileName.subSequence(0, fileName.trim().length()-4) + "_NEW.xls";
				FileUtils.copyFile(new File(strOldFileName), new File(fileName)); }
			System.out.println(">>>>>>>>>>>>>>>>>>>>>>>>  1 ");
			
			int strSheetIndex = 0;
			//get the rows count from the excel sheet
			int rowNum = fnGetExcelRowNum(fileName, sheetName, dataPointer);
			//get the columns count from the excel sheet
			int colNum = fnGetExcelColNum(fileName, sheetName, columnName);
			
			System.out.println("Row, column " + rowNum + " -- " + colNum);
			
			//defining the file input stream
			FileInputStream fileInputStream = new FileInputStream(new File(fileName));
	
			//Read the spreadsheet that needs to be updated 
			HSSFWorkbook strWorkBook = new HSSFWorkbook(fileInputStream);
			
			for(int temp=0; temp < strWorkBook.getNumberOfSheets(); temp++)
				if (strWorkBook.getSheetName(temp).trim().toLowerCase().contentEquals(sheetName.trim().toLowerCase()) == true)
					strSheetIndex = temp;
	
			//Access the workbook 
			HSSFSheet strworksheet = strWorkBook.getSheetAt(0);
			System.out.println(">>>>>>>>>>>>>>>>>>>>>>>>  2 ");
			//Access the worksheet, so that we can update
			HSSFCell strcell = strworksheet.getRow(rowNum).createCell(colNum);

			System.out.println(">>>>>>>>>>>>>>>>>>>>>>>>  3 " + strcell.getStringCellValue());
	
			//update the excel cell value
			if (fnIsStringNumeric(updValue)==true) {
				//System.out.println("Setting numeric vlaue in excel...");
				HSSFCellStyle styleCell = strWorkBook.createCellStyle();
				if (updValue.matches("^-?\\d+$") == false) {
					styleCell.setDataFormat(HSSFDataFormat.getBuiltinFormat("#.##"));
					strcell.setCellStyle(styleCell);
					strcell.setCellValue(Double.parseDouble(updValue)); 
					styleCell.setBorderTop((short) 1);
					styleCell.setBorderBottom((short) 1); 
					styleCell.setBorderRight((short) 1);
					styleCell.setBorderLeft((short) 1);
					strcell.setCellStyle(styleCell); }
				else if (updValue.matches("^-?\\d+$") == true) {
					styleCell.setDataFormat(HSSFDataFormat.getBuiltinFormat("#"));
					//strcell.setCellType(strcell.CELL_TYPE_NUMERIC);
					strcell.setCellStyle(styleCell);
					strcell.setCellValue(new Integer(updValue)); 
					
					styleCell.setBorderTop((short) 1);
					styleCell.setBorderBottom((short) 1); 
					styleCell.setBorderRight((short) 1);
					styleCell.setBorderLeft((short) 1);
					strcell.setCellStyle(styleCell); } }
			else if (updValue.trim().toLowerCase().contains("sum(") || dataPointer.trim().toLowerCase().contains("total")) {
				if (updValue.trim().toLowerCase().contains("sum("))
					strcell.setCellFormula(updValue);
				else if (dataPointer.trim().toLowerCase().contains("total"))
					strcell.setCellValue(updValue);
				HSSFCellStyle styleCell = strWorkBook.createCellStyle();
				styleCell.setBorderTop((short) 1);
				styleCell.setBorderBottom((short) 1); 
				styleCell.setBorderRight((short) 1);
				styleCell.setBorderLeft((short) 1);
				strcell.setCellStyle(styleCell);
				HSSFFont strfont = strWorkBook.createFont();
				strfont.setFontHeightInPoints((short)13);
				strfont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
				styleCell.setFont(strfont); }
			else {
				strcell.setCellValue(updValue);
				HSSFCellStyle styleCell = strWorkBook.createCellStyle();
				styleCell.setBorderTop((short) 1);
				styleCell.setBorderBottom((short) 1); 
				styleCell.setBorderRight((short) 1);
				styleCell.setBorderLeft((short) 1);
				strcell.setCellStyle(styleCell); }
					
			fileInputStream.close();
			FileOutputStream fileOutputStream = new FileOutputStream(new File(fileName));
			strWorkBook.write(fileOutputStream);
			fileOutputStream.close(); 
			System.out.println(">>>>>>>>>>>>>>>>>>>>>>>> 4 ");
			
			if (strOldFileName.trim().toLowerCase().contains(".xls")) {
				FileUtils.copyFile(new File(fileName), new File(strOldFileName));
				File tempFile = new File(fileName);
				if (tempFile.exists())
					tempFile.delete(); } }
		catch (Exception exp) {
			System.out.println("[ERROR]    Error encountered in updating the data/ result file - " + fileName + ", Sheet Name : " + sheetName + 
					", DataPointer - " + dataPointer + ", Column Name - " + columnName);
			System.out.println(exp.getMessage());
			intreturnvalue = 1;}
		return intreturnvalue;	}
	
//function to update the excel sheet cell data=============================================================================================================================
	@SuppressWarnings("resource")
	public synchronized int fnUpdateExcelCellData(String fileName, String sheetName, String dataPointer, String columnName, String updValue) throws Exception {
		int intreturnvalue = 0;
		try {
			//fnAppSync(2);
			String strOldFileName = fileName;
			int strSheetIndex = 0;

			//get the rows count from the excel sheet
			int rowNum = fnGetExcelRowNum(fileName, sheetName, dataPointer);

			//get the columns count from the excel sheet
			int colNum = fnGetExcelColNum(fileName, sheetName, columnName);
			
			if ((rowNum > 0) && (colNum > 0)) {
				//defining the file input stream
				FileInputStream fileInputStream = new FileInputStream(new File(fileName));
		
				//Read the spreadsheet that needs to be updated 
				HSSFWorkbook strWorkBook = new HSSFWorkbook(fileInputStream);
				
				for(int temp=0; temp < strWorkBook.getNumberOfSheets(); temp++)
					if (strWorkBook.getSheetName(temp).trim().toLowerCase().contentEquals(sheetName.trim().toLowerCase()) == true)
						strSheetIndex = temp;
	
				//Access the workbook
				HSSFSheet strworksheet = strWorkBook.getSheet(sheetName.trim());
				
				//Access the worksheet, so that we can update
				Cell strcell = strworksheet.getRow(rowNum).createCell(colNum);
				strcell.setCellValue(updValue.trim());
				fileInputStream.close();
				FileOutputStream fileOutputStream = new FileOutputStream(new File(fileName));
				strWorkBook.write(fileOutputStream);
				fileOutputStream.close(); 
				fnAppSync(2); }
			else {
				System.out.println("[ERROR]    Incorrect data pointer/ row not found in data file - " + dataPointer + ", " + columnName); }
		}
		catch (Exception exp) {
			System.out.println("[ERROR]    Error updating the data/ result file - " + fileName + ", Sheet Name : " + sheetName + 
					", DataPointer - " + dataPointer + ", Column Name - " + columnName + ", Column Value : " + updValue);
			System.out.println("[ERROR]    " + exp.getStackTrace());
			System.out.println("[ERROR]    " + exp.getLocalizedMessage());
			System.out.println("[ERROR]    " + exp.getCause());

			try{
				fnAppSync(3);
				System.out.println("[INFO]    One more try for updating the data/ result file - " + fileName);
				int strSheetIndex = 0;

				//get the rows count from the excel sheet
				int rowNum = fnGetExcelRowNum(fileName, sheetName, dataPointer);

				//get the columns count from the excel sheet
				int colNum = fnGetExcelColNum(fileName, sheetName, columnName);

				if ((rowNum > 0) && (colNum > 0)) {
					//defining the file input stream
					FileInputStream fileInputStream = new FileInputStream(new File(fileName));
			
					//Read the spreadsheet that needs to be updated 
					HSSFWorkbook strWorkBook = new HSSFWorkbook(fileInputStream);
					
					for(int temp=0; temp < strWorkBook.getNumberOfSheets(); temp++)
						if (strWorkBook.getSheetName(temp).trim().toLowerCase().contentEquals(sheetName.trim().toLowerCase()) == true)
							strSheetIndex = temp;
		
					//Access the workbook 
					HSSFSheet strworksheet = strWorkBook.getSheet(sheetName.trim());
					
					//Access the worksheet, so that we can update
					Cell strcell = strworksheet.getRow(rowNum).createCell(colNum);
					strcell.setCellValue(updValue.trim());
					fileInputStream.close();
					FileOutputStream fileOutputStream = new FileOutputStream(new File(fileName));
					strWorkBook.write(fileOutputStream);
					fileOutputStream.close(); 
					fnAppSync(2); }
				else {
					System.out.println("[ERROR]    Incorrect data pointer/ row not found in data file - " + dataPointer + ", " + columnName); } }
			catch (Exception exp1){	
				intreturnvalue=1; }	}
		return intreturnvalue; }
		
	//function to verify the object is displayed or not=============================================================================================================================
	public int fnGetObjectDisplayed(String objAttribute, int intWaitTime) throws Throwable {
		int strFound = 1;
		for(int intTemp = 0; intTemp < intWaitTime; intTemp++)
		{ 	fnAppSync(pageObjects.gnWaitTime);
			if (wDriver.findElements(By.xpath(objAttribute)).size() == 1) strFound = 0;
			if (strFound == 0) break; 		}
		return strFound; }
	
	//Function to validate the object displayed or not and then enter the value====================================================================================================
	public void fnObjEnterValueonFirstDisplay(String strObject, String strVal) throws Throwable {

		java.util.List<WebElement> strList = wDriver.findElements(By.xpath(strObject));
		fnAppSync(pageObjects.gnWaitTime);
		for(int intVal=0; intVal < strList.size(); intVal++){
			java.util.List<WebElement> strList1 = wDriver.findElements(By.xpath(strObject));
			if (strList1.get(intVal).isDisplayed() == true) {
					fnAppSync(pageObjects.gnWaitTime);
					strList1.get(intVal).click();
					fnAppSync(pageObjects.gnWaitTime);
					strList1.get(intVal).clear();					
					strList1.get(intVal).sendKeys(strVal);
					fnAppSync(pageObjects.gnWaitTime);
					strList1.get(intVal).sendKeys(Keys.TAB);
					fnAppSync(pageObjects.gnWaitTime);
					break; } } }

	//Function to set the excel cell background color=============================================================================================================================
	@SuppressWarnings({ "deprecation", "resource" })
	public synchronized void fnExcelCellColor(String strFileName, String strSheetName) throws Throwable {
		int strSheetIndex = 0;
		
		FileInputStream fsIP= new FileInputStream(new File(strFileName)); 
		//Read the spreadsheet that needs to be updated 
		HSSFWorkbook wb = new HSSFWorkbook(fsIP);
		
		for(int temp=0; temp<wb.getNumberOfSheets(); temp++)
			if (wb.getSheetName(temp).trim().toLowerCase().contentEquals(strSheetName.trim().toLowerCase()) == true)
				strSheetIndex = temp;
		
		//Access the workbook or sheet name with index
		HSSFSheet worksheet = wb.getSheetAt(strSheetIndex);
		
		Row row = worksheet.getRow(2);
		CellStyle style = wb.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    Cell cell1 = row.getCell(4);
	    cell1.setCellValue("COLOR TEST");
	    cell1.setCellStyle(style);
		
	    //Close the InputStream 
  		FileOutputStream output_file =new FileOutputStream(new File(strFileName));
  		
  		//Open FileOutputStream to write updates 
  		wb.write(output_file); //write changes 
  		output_file.close(); 
  		fsIP.close(); }
	
	//Function to delete all the excel sheet from the excel file except the given sheet name===========================================================================
	@SuppressWarnings("resource")
	public synchronized void fnExcelDelSheetsExcept(String strFileName, String strSheetName) throws Throwable {
				
		//defining the file input stream
		FileInputStream fileInputStream = new FileInputStream(new File(strFileName));
		//Read the spreadsheet that needs to be updated 
		HSSFWorkbook strWorkBook = new HSSFWorkbook(fileInputStream);
		//String strSheetExempted = "EQO_Asian_ResetTabValues;ComplexSwap_OtherFee;ComplexSwap_PeriodData;cds_defquery;CDS_OtherFees;Swaption_OtherFee;Callable_OtherFee;Callable_UnderLying_NextCall";
		if (strSheetName.trim().toLowerCase().contentEquals("testcase_loan_bulkkill") || strSheetName.trim().toLowerCase().contentEquals("testcase_deposit_bulkkill"))
			strSheetExempted = pageObjects.strBulkData;
		
		for(int temp=strWorkBook.getNumberOfSheets()-1; temp >= 0; temp--)
			if (strWorkBook.getSheetName(temp).trim().toLowerCase().contentEquals(strSheetName.trim().toLowerCase()) == false)
				if (strSheetExempted.trim().toLowerCase().contains(strWorkBook.getSheetName(temp).trim().toLowerCase()) == false){
					int strDelSht = 0;

					if (strWorkBook.getSheetName(temp).trim().toLowerCase().contains("eqo_asian_resettabvalues") == true)
						strDelSht = 1;
					else if (strWorkBook.getSheetName(temp).trim().toLowerCase().contains("cds_defquery") == true)
						strDelSht = 1;

					if (strDelSht == 0)
						strWorkBook.removeSheetAt(temp);
					else if (strWorkBook.getSheetName(temp).trim().toLowerCase().contentEquals(strSheetExempted.toLowerCase()) == false)  
						strWorkBook.removeSheetAt(temp); }

		fileInputStream.close();
		FileOutputStream fileOutputStream = new FileOutputStream(new File(strFileName));
		strWorkBook.write(fileOutputStream);
		fileOutputStream.close(); }
	
	//Function to wait for the object to be displayed or loaded onto the screen, returns 1 - not displayed, 0 - displayed==================================================
	public int fnWaitForObjectDisplay(String strObject) throws Throwable {
		int strDisp = 1;
		for(int intVal=1; intVal < 60; intVal++)
		{	fnAppSync(pageObjects.gnWaitTime);
			if (wDriver.findElements(By.xpath(strObject)).size() > 0)
				strDisp = 0;
			if (strDisp == 0) break; }
		return strDisp; }

	//Function to clear the dasboard=============================================================================================================================
	public void fnClearDPMDashboard() throws Throwable {
		try{
			if (fnObjectExists(objMenuLnk, 1) != 0) {
				fnObjectClick(objMenuLnk);
				fnAppSync(1);
				if (fnObjectExists(objClrDshBrd, 1) != 0)
					fnObjectClick(objClrDshBrd);
				fnAppSync(1);
				System.out.println("[INFO]   Successfully cleared the dashboard..... .....");
				//step to click on the notification icon
				if (fnObjectExists(objPXVLabel, 1) != 0)
					fnObjectClick(objPXVLabel); } }
		catch(Exception exp)
		{ System.out.println("[ERROR]   Unable to click on the Menu link......");
		  pageObjects.strLogoutStatus = true;}	}
		
	//Function to prepare the PDF doc with the test case screen shots==============================================================================================================
	public void fnPDFScreenReport(String dataPointer) throws Throwable{
		System.out.println("[INFO]   Screens report PDF - Generated");
		if (pageObjects.strGenPDF == true) {
			Document document = new Document(PageSize.A4);
			Font font;
			PdfWriter pdfWrite = PdfWriter.getInstance(document, new FileOutputStream(strScreenPath.trim() + dataPointer.trim() + ".pdf"));
			pageObjects.intExecCount = pageObjects.intExecCount + 1;
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
			System.out.println("[INFO]    ************ Total Wait Time : " + pageObjects.totWaitTime + " Seconds");
			pageObjects.totWaitTime = 0;
			document.close(); } 
		
			if (onErrorLogout == true){
				onErrorLogout = false;
				loginLogout fnLoginModule = new loginLogout();
				fnLoginModule.dpmLogout(); } }
	
	//function to get the current date and time =========================================================================================================================
	public String fnGetCurrentDate(String strFormat) throws Throwable, Exception{
		String strYear, strMonth, strDay, strRet = null;
		
		GregorianCalendar date = new GregorianCalendar();
			strYear = String.valueOf(date.get(Calendar.YEAR));
		
		if ((date.get(Calendar.MONTH) + 1) < 10)
			strMonth = "0" + String.valueOf(date.get(Calendar.MONTH)+1);
		else
			strMonth = String.valueOf((date.get(Calendar.MONTH))+1);
		
		if (date.get(Calendar.DAY_OF_MONTH) < 10)
			strDay = "0" + String.valueOf(date.get(Calendar.DAY_OF_MONTH));
		else
			strDay = String.valueOf(date.get(Calendar.DAY_OF_MONTH));

		if (strFormat.trim().toLowerCase().contentEquals("yyyymmdd") == true)
			strRet = strYear + strMonth + strDay;
		else if (strFormat.trim().toLowerCase().contentEquals("yyyy/mm/dd") == true)
			strRet =  strYear + "/" + strMonth + "/" + strDay;
		else if (strFormat.trim().toLowerCase().contentEquals("mm-dd-yyyy") == true)
			strRet =  strDay + "-" + strMonth + "-" + strYear;
		else if (strFormat.trim().toLowerCase().contentEquals("yyyy-mm-dd") == true)
			strRet =  strYear + "-" + strMonth + "-" + strDay;
		else if (strFormat.trim().toLowerCase().contentEquals("dd-mm-yyyy") == true)
			strRet =  strDay + "-" + strMonth + "-" + strYear;
		
		
		return strRet; 	}
	
	//function to get the current time ==========================================================================================================================================
	public String fnGetCurrentTime(){
		String strCurrTime1 = new SimpleDateFormat("HH.mm").format(new Date());
		return strCurrTime1; }

	//function to search for the todays folder exists or not, if no create one ==================================================================================================
	public void fnCreateResultsFolder() throws Throwable, Throwable {
			if (pageObjects.strFeature.trim().toLowerCase().contains("moneymarket_loan") || pageObjects.strFeature.trim().toLowerCase().contains("moneymarket_deposit"))
				pageObjects.strScreenPath = pageObjects.strScreenPath + "\\\\MoneyMarket\\\\" + pageObjects.strFunctionality + "\\\\" + pageObjects.strReleaseVersion + "\\\\" + pageObjects.strGVDate;
			else if (pageObjects.strFeature.trim().toLowerCase().contains("complexswap_feature"))
				pageObjects.strScreenPath = pageObjects.strScreenPath + "\\\\Complex_Swap_Feature\\\\" + pageObjects.strFunctionality + "\\\\" + pageObjects.strReleaseVersion + "\\\\" + pageObjects.strGVDate;
			else if (pageObjects.strFeature.trim().toLowerCase().contains("cds_feature"))
				pageObjects.strScreenPath = pageObjects.strScreenPath + "\\\\CDS_Feature\\\\" + pageObjects.strFunctionality + "\\\\" + pageObjects.strReleaseVersion + "\\\\" + pageObjects.strGVDate;
			else
				pageObjects.strScreenPath = pageObjects.strScreenPath + "\\\\" + pageObjects.strFeature + "\\\\" + pageObjects.strFunctionality +"\\\\" +pageObjects.strReleaseVersion + "\\\\" + pageObjects.strGVDate;
			
			// step to creat the rlease -> date wise -> count folder
			File folder = new File(pageObjects.strScreenPath);
			File[] foldList = folder.listFiles();
			int intFldCount=1;

			if (folder.exists() == false){
				try { folder.mkdir(); }
				catch(Exception exp){ System.out.println("[ERROR]    Failed to create the results folder, error : " + exp.getMessage()); } }
			else
				intFldCount = (foldList.length) + 1;
			fnAppSync(1);
			
			pageObjects.strScreenPath = pageObjects.strScreenPath + "\\\\" + pageObjects.strGVDate + "_" + intFldCount + "_" + new SimpleDateFormat("HH.mm").format(new Date());
			System.out.println("[INFO]   Folder to create ... " + pageObjects.strScreenPath);
			File newFolder = new File(pageObjects.strScreenPath);
			newFolder.mkdirs();
			pageObjects.strScreenPath = pageObjects.strScreenPath + "\\\\";
			pageObjects.strHTMLPath = pageObjects.strScreenPath;
			System.out.println("[INFO]   Created Folder for HTML PATH " + pageObjects.strHTMLPath);
			try{
				System.out.println("[INFO]   strDataSheetPath -- " + pageObjects.strDataSheetPath);
				System.out.println("[INFO]   strDataFileName -- " + pageObjects.strDataFileName);
				System.out.println("[INFO]   strConfigFileName -- " + pageObjects.strConfigFileName);
				System.out.println("[INFO]   strHTMLPath -- " + pageObjects.strHTMLPath);
				System.out.println("[INFO]   strConfigFilepath -- " + pageObjects.strConfigFilepath);
				
				FileUtils.copyFile(new File(pageObjects.strDataSheetPath), new File(pageObjects.strHTMLPath + pageObjects.strDataFileName));
				System.out.println("[INFO]   Copied the data file - " + pageObjects.strDataSheetPath + ", to the path - " + pageObjects.strHTMLPath);
				FileUtils.copyFile(new File(pageObjects.strConfigFilepath), new File(pageObjects.strHTMLPath + pageObjects.strConfigFileName));
				
				folder = new File(strDefaultChromeDriverPathName);
				System.out.println(folder.exists());
				if (folder.exists()==false) {
					folder.mkdirs(); }
				FileUtils.copyFile(new File(strDefaultChromeDriverSrcPathName + strDefaultChromeFileName), new File(pageObjects.strHTMLPath + "chromedriver.exe"));
				
				//FileUtils.copyFile(new File(pageObjects.strDefaultChromeDriverPathName + strDefaultChromeFileName), new File(pageObjects.strHTMLPath + strDefaultChromeFileName));
				pageObjects.strChromeDriverPath = strDefaultChromeDriverPathName + strDefaultChromeFileName;
				fnAppSync(5);

				//Step to create a text file for updating to the PXV-CONFIG FILE ==============================================================================================
				fnWriteTxtFile(strDefaultConfigValuesPathName + pageObjects.strFunctionality + ".txt", pageObjects.strFunctionality + ":" + pageObjects.strHTMLPath + pageObjects.strDataFileName);
				System.out.println("[INFO]   Created the CONFIG Update file for process - " + strDefaultConfigValuesPathName + pageObjects.strFunctionality + ".txt");
			}
			catch (Exception exp) { System.out.println("[ERROR}   - " +    exp.getMessage());}
			pageObjects.strResultsPath = pageObjects.strHTMLPath;
			pageObjects.strConfigFilepath = pageObjects.strHTMLPath + pageObjects.strConfigFileName;
			pageObjects.strDataSheetPath = pageObjects.strHTMLPath + pageObjects.strDataFileName; 
			pageObjects.strScreenPath1 = pageObjects.strScreenPath; 
			
			//Step to delete the files/ folders from the temp
			//fnDeleteTmpFilesFolder(System.getProperty("java.io.tmpdir"), ""); 
			}

	//function set the pre-configuration parameters from the input file ====================================================================================================
	public void fnLoadConfigurationParameters() throws Throwable{
			InetAddress IP=InetAddress.getLocalHost();
			System.out.println("[INFO]    Running on machine - IP == " + IP.getHostAddress());
		
			String strTempFilePath = pageObjects.strDefaultPath + pageObjects.strDefaultConfigName;
			String strVal="";
			File tempFile = new File(strTempFilePath);
			if (tempFile.exists() == false)
				pageObjects.strDefaultPath = pageObjects.strDefaultPath1;
		
			strDataSheetPath = strDefaultPath;
			strConfigFilepath = strDefaultPath;
			strConfigFileName = strDefaultConfigName;
			pageObjects.strDealAmend = false;
			pageObjects.strComplexFlip = false;
			pageObjects.strDBCompareVal = false;
			pageObjects.strValSwaptionDetails = false;
			pageObjects.strValUnderlyingSwapDetails = false;
			pageObjects.strCSSaveNewTemplate = false;
			pageObjects.strSaveAs = false;
			pageObjects.intExecCount = 0;
			pageObjects.strScreenFileCnt = 0;
			pageObjects.strTrdActRstLayout = true;
			pageObjects.strComplexFlip = false;
			strDataFileName = null;
			pageObjects.strScreenPath = "";
			pageObjects.strMultipleDealIds = "";
			pageObjects.strDataFileName = "peakxv_" + pageObjects.strProductName + "_TestData.xls";
			pageObjects.strExecutionFlag = fnGetExcelData(strTempFilePath, "DPMConfig", "strExecutionFlag", "ConfigValue").trim().toLowerCase();
			pageObjects.strSTPExecutionFlag = fnGetExcelData(strTempFilePath, "DPMConfig", "strSTPExecutionFlag", "ConfigValue").trim().toLowerCase();
			pageObjects.strSTPPXVExecutionFlag = fnGetExcelData(strTempFilePath, "DPMConfig", "strSTPPXVExecutionFlag", "ConfigValue").trim().toLowerCase();

			System.out.println("[INFO]   Functionality  - " + pageObjects.strFunctionality);
			if ((pageObjects.strExecutionFlag.contains("y")==false) && (pageObjects.strInputSheetName.trim().toLowerCase().contains("stp_")==false)) {
				Character strLastVal = pageObjects.strExecutionFlag.charAt(pageObjects.strExecutionFlag.length()-1);
				
				if (Character.isDigit(strLastVal))
					strVal = pageObjects.strFunctionality + strLastVal;
				else
					strVal = pageObjects.strFunctionality;

				strVal = fnGetExcelData(strTempFilePath, "DPMConfig", strVal, "ConfigValue").trim();
				if (strVal.trim().length()>0) {
					String strVal1 = new StringBuffer(strVal).reverse().toString();
					int intLoc = strVal1.indexOf("\\");				
					String strVal2 = strVal1.substring(intLoc+1, strVal1.length());
					String strVal3 = new StringBuffer(strVal2).reverse().toString() + "\\";
					
					strDefaultPath = strVal3;
					System.out.println("Data Sheet path : " + strVal3);
					pageObjects.strDataSheetPath = strVal;	}

				System.out.println("[INFO]   PXV Data File Name - " + strDefaultPath + pageObjects.strDataFileName); }
			else if ((pageObjects.strSTPExecutionFlag.trim().toLowerCase().contentEquals("y")==true) && (pageObjects.strInputSheetName.trim().toLowerCase().contains("stp_")==true)) {
				
				System.out.println("[INFO]   STP Data File Name - " + strDefaultPath + pageObjects.strDataFileName);
				System.out.println("[INFO]   STP Data CONFIG File Name - " + strDefaultPath + pageObjects.strConfigFileName); }
			else if ((pageObjects.strSTPExecutionFlag.trim().toLowerCase().contentEquals("executed")==true) && (pageObjects.strInputSheetName.trim().toLowerCase().contains("stp_")==true)) {
				String strTempVal = fnGetExcelData(strTempFilePath, "DPMConfig", pageObjects.strInputSheetName, "ConfigValue").trim().toLowerCase();
				
				String strVal1 = new StringBuffer(strTempVal).reverse().toString();
				int intLoc = strVal1.indexOf("\\");
				String strVal2 = strVal1.substring(intLoc+1, strVal1.length());
				String strVal3 = new StringBuffer(strVal2).reverse().toString() + "\\";

				strDataFilePath = strVal3;
				System.out.println("[INFO]   Data Sheet path : " + strVal3);
				pageObjects.strDataSheetPath = strVal3;
				System.out.println("[INFO]   STP-PXV Data File Name - " + pageObjects.strDataFileName); }
			
			System.out.println("[INFO]   STP-PXV Data CONFIG File Name - " + strDefaultPath + pageObjects.strConfigFileName);
			pageObjects.strSTPFilePath = pageObjects.strDefaultPath + pageObjects.strConfigFileName;
			strTempFilePath = pageObjects.strDefaultPath + pageObjects.strDefaultConfigName;
			pageObjects.strScreenPath = fnGetExcelData(strTempFilePath, "DPMConfig", "ImgScreenPath", "ConfigValue").trim();
			pageObjects.strChromeDriverPath = fnGetExcelData(strTempFilePath, "DPMConfig", "ChromeDriverPath", "ConfigValue").trim();
			pageObjects.strReleaseVersion = fnGetExcelData(strTempFilePath, "DPMConfig", "ReleaseVersion", "ConfigValue").trim();
			pageObjects.strTriggerEmail = fnGetExcelData(strTempFilePath, "DPMConfig", "eMailTrigger", "ConfigValue").trim();
			pageObjects.pgWaitTime = Integer.valueOf(fnGetExcelData(strTempFilePath, "DPMConfig", "pageWaitTime", "ConfigValue"));
			pageObjects.objWaitTime = Integer.parseInt(fnGetExcelData(strTempFilePath, "DPMConfig", "objWaitTime", "ConfigValue"));
			pageObjects.gnWaitTime = Integer.parseInt(fnGetExcelData(strTempFilePath, "DPMConfig", "gnWaitTime", "ConfigValue"));
			pageObjects.strConfluenceSheetPath = fnGetExcelData(strTempFilePath, "DPMConfig", "stConfluencePath", "ConfigValue").trim();
			pageObjects.strTradeDate = fnGetExcelData(strTempFilePath, "DPMConfig", "TradeDate", "ConfigValue").trim();
			pageObjects.strInitialStrikeDate = fnGetExcelData(strTempFilePath, "DPMConfig", "InitialStrikeDate", "ConfigValue").trim();
			pageObjects.strExpiryDate = fnGetExcelData(strTempFilePath, "DPMConfig", "ExpiryDate", "ConfigValue").trim();
			pageObjects.strSettlementDate = fnGetExcelData(strTempFilePath, "DPMConfig", "SettlementDate", "ConfigValue").trim();
			pageObjects.strExportedFilePath = System.getProperty("user.home") + "\\Downloads\\";

			if (pageObjects.strAppName.trim().toLowerCase().contentEquals("pxv")){
				pageObjects.strLoginSuperUserID = fnGetExcelData(strTempFilePath, "DPMConfig", "strPXVSuperUser", "ConfigValue").trim(); }
			if (pageObjects.strAppName.trim().toLowerCase().contentEquals("markit")){
				pageObjects.strLoginSuperUserID = fnGetExcelData(strTempFilePath, "DPMConfig", "strMARKITWIRESuperUser", "ConfigValue").trim(); }
			pageObjects.strLoginUserID = pageObjects.strLoginSuperUserID;
			
			if (fnGetExcelData(strTempFilePath, "DPMConfig", "DoNotUseDealMenu", "ConfigValue").trim().toLowerCase().equalsIgnoreCase("true"))
				pageObjects.strUseDealMenu = false;
			else
				pageObjects.strUseDealMenu = true;
	
			pageObjects.strGVDate = fnGetCurrentDate("yyyymmdd");
			pageObjects.strGVTime = fnGetCurrentTime();
			
			if (pageObjects.strExecutionFlag.contentEquals("fail")==false)
				pageObjects.strDataSheetPath = pageObjects.strDataSheetPath + pageObjects.strDataFileName;
			
			if (pageObjects.strDataSheetPath.trim().toLowerCase().contains(pageObjects.strDataFileName.trim().toLowerCase())==false)
				pageObjects.strDataSheetPath = pageObjects.strDataSheetPath + pageObjects.strDataFileName;
			
			pageObjects.strConfigFilepath = pageObjects.strConfigFilepath + pageObjects.strConfigFileName;
			
			System.out.println("[INFO]   strDataSheetPath - " + pageObjects.strDataSheetPath);
			System.out.println("[INFO]   email - " + pageObjects.strTriggerEmail);
			System.out.println("[INFO]   release ver - " + pageObjects.strReleaseVersion);
			System.out.println("[INFO]   Execution Flag - " + pageObjects.strExecutionFlag);
			System.out.println("[INFO]   Execution date - " + pageObjects.strGVDate);
			System.out.println("[INFO]   Default User/ SuperUser ID - " + pageObjects.strLoginUserID);
			
			//Steps to set the Execution Flag to Fail in the data sheets =============================================
			if (pageObjects.strExecutionFlag.trim().toLowerCase().contentEquals("fail")) {
				String strVal3 = fnGetExcelData(strDefaultConfigFilePathName, "DPMConfig", pageObjects.strFunctionality, "ConfigValue").trim();
				if (strVal3.trim().length()>0) {
					int intRowCount  = fnGetExcelRowCount(pageObjects.strDataSheetPath, pageObjects.strInputSheetName);
					for (int intLoop = 1; intLoop < intRowCount; intLoop++){
						strVal = fnGetExcelCellData(pageObjects.strDataSheetPath, pageObjects.strInputSheetName, intLoop, "Test_Status");
						String strVal1 = fnGetExcelCellData(pageObjects.strDataSheetPath, pageObjects.strInputSheetName, intLoop, "TestCase");
						if (strVal.trim().toLowerCase().contentEquals("fail")) {
							fnUpdateExcelCellData(pageObjects.strDataSheetPath, pageObjects.strInputSheetName, strVal1, "ExecutionFlag", "Y"); 
							fnAppSync(5); }	} } }
			else if (pageObjects.strExecutionFlag.trim().toLowerCase().contentEquals("pass")) {
				String strVal3 = fnGetExcelData(strDefaultConfigFilePathName, "DPMConfig", pageObjects.strFunctionality, "ConfigValue").trim();
				if (strVal3.trim().length()>0) {
					int intRowCount  = fnGetExcelRowCount(pageObjects.strDataSheetPath, pageObjects.strInputSheetName);
					for (int intLoop = 1; intLoop < intRowCount; intLoop++){
						strVal = fnGetExcelCellData(pageObjects.strDataSheetPath, pageObjects.strInputSheetName, intLoop, "Test_Status");
						String strVal1 = fnGetExcelCellData(pageObjects.strDataSheetPath, pageObjects.strInputSheetName, intLoop, "TestCase");
						if (strVal.trim().toLowerCase().contentEquals("fail")) {
							fnUpdateExcelCellData(pageObjects.strDataSheetPath, pageObjects.strInputSheetName, strVal1, "ExecutionFlag", "Y"); 
							fnAppSync(5); }	} } }
			else if (pageObjects.strExecutionFlag.trim().toLowerCase().contentEquals("all")) {
				String strVal3 = fnGetExcelData(strDefaultConfigFilePathName, "DPMConfig", pageObjects.strFunctionality, "ConfigValue").trim();
				if (strVal3.trim().length()>0) {
					int intRowCount  = fnGetExcelRowCount(pageObjects.strDataSheetPath, pageObjects.strInputSheetName);
					for (int intLoop = 1; intLoop < intRowCount; intLoop++){
						strVal = fnGetExcelCellData(pageObjects.strDataSheetPath, pageObjects.strInputSheetName, intLoop, "Test_Status");
						String strVal1 = fnGetExcelCellData(pageObjects.strDataSheetPath, pageObjects.strInputSheetName, intLoop, "TestCase");
						if ((strVal.trim().toLowerCase().contentEquals("fail")) || (strVal.trim().toLowerCase().contentEquals("pass"))) {
							fnUpdateExcelCellData(pageObjects.strDataSheetPath, pageObjects.strInputSheetName, strVal1, "ExecutionFlag", "Y"); 
							fnAppSync(5); }	} } } }

	//function to generate the HTML report for test execution =================================================================================================================
	public synchronized void fnGenerateHTMLHeaderReport(String strHTMLFilePath) throws Throwable, Throwable {

		//Steps to create the HTML report header section
		FileWriter strHTMLFile;
		String strFoldPath = "\\" + (pageObjects.strScreenPath).replace("\\\\", "\\");
		System.out.println("[INFO]   Folder path -- -- -- -- -- -- " + strFoldPath);

		String strCurrTime1 = new SimpleDateFormat("HH.mm").format(new Date());
		String strHeader = "<html><head> <style> a:link { color: black; } a:visited { color: black; } a:hover { color: Blue; } a:active { color: blue; } </style> </head>"
				+ "<font face='verdana'><h4 align='center'> PEAK XV (DPM) Automation Execution Report - Environment : " + pageObjects.strEnvironment + ", " + pageObjects.strReleaseVersion +"</h4><hr><td>"
						+ "<h6 align='left'><font color='blue' SIZE=2> URL : " + pageObjects.strEnvironmentURL + ",    User ID. " + pageObjects.strLoginUserID + ", Exec.Date: " + fnGetCurrentDate("yyyy/mm/dd") + ", Time : " + strCurrTime1 + "</font> </h6></td>"
								//+ "<td><h6><font color='blue'  SIZE=2> Execution - Date : " + fnGetCurrentDate("yyyy/mm/dd") + ", Time : " + strCurrTime1 + "</font></h6>"
										+ "<h6 align ='left'><font color='blue' SIZE=2>"
										+ "<a href=" + "\"" + strFoldPath + "\"" +  "onclick = open(this.href, this.target, 'width=600, height=450, top=200, left=250'); return false; target=_blank> Click to open Results Folder.. ..</a>"
										+ "</font></h6> "
										+ "<h6><font color='black' SIZE=2> Product : " + pageObjects.strProductName + " </h6> "
										+ "<h6><font color='black' SIZE=2> Feature : " + pageObjects.strFeatureName + " </h6> </td>"
										+ "<TABLE BORDER=2 BGCOLOR=#ffff00>"
										+ "<TR><TH><FONT COLOR=BLACK FACE='Geneva, Arial' SIZE=2>Use_Case/Story_Name_________</FONT></TH>"
										+ "<TH><FONT COLOR=BLACK FACE='Geneva, Arial' SIZE=2>Status</TH>"
										+ "<TH><FONT COLOR=BLACK FACE='Geneva, Arial' SIZE=2>PEAK XV_DealID</TH>"
										+ "<TH><FONT COLOR=BLACK FACE='Geneva, Arial' SIZE=2>TimeStamp</TH>"
										+ "<TH><FONT COLOR=BLACK FACE='Geneva, Arial' SIZE=2>Remarks______________________________________________</TH></TR>";
				
		pageObjects.strHTMLName = "PXV_AutRep_" + pageObjects.strFunctionality + "_" + fnGetCurrentDate("yyyymmdd") + ".HTML";

		//step to set the output stream to a log file
		String strTxtFile = "PXV_Log_" + pageObjects.strProductName + "_" + fnGetCurrentDate("yyyymmdd") + ".TXT";
		System.setOut(new PrintStream(new BufferedOutputStream(new FileOutputStream(pageObjects.strHTMLPath + strTxtFile)), true));	

		File strFile = new File(pageObjects.strHTMLPath + pageObjects.strHTMLName);
		if (strFile.exists() == false) {
			strFile.createNewFile();
			strHTMLFile = new FileWriter(pageObjects.strHTMLPath + pageObjects.strHTMLName, true);
			strHTMLFile.write(strHeader);
			strHTMLFile.close(); }
		
		//Steps to delete all the other sheets except the required sheet
		fnExcelDelSheetsExcept(pageObjects.strDataSheetPath, pageObjects.strInputSheetName);
		
		//Step to create a PXV_CONFIG SHEET details txt file ===============================================================================================
		if (pageObjects.strWriteConfig == true)
			fnWriteTxtFile(strDefaultConfigValuesPathName + pageObjects.strFunctionality + ".txt", pageObjects.strFunctionality + ":" + pageObjects.strHTMLPath + pageObjects.strDataFileName); }

	//function to generate the HTML results row one by one
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

		pageObjects.strPDFPath =  strTempPath1 + strTC + ".PDF ";
		if (pageObjects.strAddUser == true) {
			File file1 = new File(pageObjects.strPDFPath);
			if (file1.exists()){
				String strTC1 = strTempPath1 + strTC + "_User_" + pageObjects.strLoginUserID + ".PDF";
				File file2 = new File(strTC1);
				if (file2.exists()==false) {
					if (file1.renameTo(file2)){
						System.out.println("[INFO]   File renamed with user id - " + strTC1); 
						strTC = strTC + "_User_" + pageObjects.strLoginUserID + ".PDF";
						pageObjects.strPDFPath =  strTC1; }
					else { System.out.println("[INFO]   Failed to rename the file - " + strTC); } } } }

		if (strStatus.trim().toLowerCase().contentEquals("pass") == true)
			strColor = "#006633"; // green color code
		else if (strStatus.trim().toLowerCase().contentEquals("fail") == true)
			strColor = "#D80000"; //red color code
		else if(strStatus.trim().toLowerCase().contentEquals("pass&fail") == true)
			strColor = "#FF69B4"; // hotpink color code

		String strCurrTime1 = new SimpleDateFormat("HH.mm").format(new Date());
		//String strRows = "<TR  style='background:silver'" + ";text-align='left'>";
		//String strRows = "<TR  style='background:#D8D8D8'" + ";text-align='left'>";
		String strRows = "<TR  style='background:#f2f2f2'" + ";text-align='left'>";

		//strRows = strRows + "<TH BGCOLOR = '" + strColor + "' ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=2> <a href=" + "\"" + pageObjects.strPDFPath + "\"" + "onclick = open(this.href, this.target, 'width=600, height=450, top=200, left=250'); return false; target=_blank>" + strTC + "</a> </FONT></TH>";
		//strRows = strRows + "<TH BGCOLOR = '" + strColor + "' ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=2>" + strStatus + " </FONT></TH>";
		if(pageObjects.strGVTCTemp.contentEquals(strTC) ){
			strRows = strRows + "<TH BGCOLOR = '' ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=1> <a href=" + "\"" + pageObjects.strPDFPath + "\"" + "onclick = open(this.href, this.target, 'width=600, height=450, top=200, left=250'); return false; target=_blank><font color='" + strColor + "'>" + "" + "</font></a> </FONT></TH>";
		}else{
			strRows = strRows + "<TH BGCOLOR = '' ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=1> <a href=" + "\"" + pageObjects.strPDFPath + "\"" + "onclick = open(this.href, this.target, 'width=600, height=450, top=200, left=250'); return false; target=_blank><font color='" + strColor + "'>" + strTC + "</font></a> </FONT></TH>";
			pageObjects.strGVTCTemp = strTC;
		}
		strRows = strRows + "<TH BGCOLOR = '' ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=1><font color='" + strColor + "'>" + strStatus + "</font> </FONT></TH>";

		strRows = strRows + "<TH ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=1>" + strDealID + " </FONT></TH>";
		strRows = strRows + "<TH ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=1>" + strCurrTime1 + " </FONT></TH>";
		strRows = strRows + "<TH STYLE='WORD-WRAP:BREAK-WORD' ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=1>" + strRemarks + " </FONT></TH></TR>";

		//Steps to update the test failure details in ConfluenceDashBoardResults.xls - FailedTests sheet
		if (fnGetExcelData(pageObjects.strConfigFilepath, "DPMConfig", "UpdateConfluence", "ConfigValue").trim().toLowerCase().contentEquals("true"))
			if (strStatus.trim().toLowerCase().toLowerCase().contentEquals("fail")) {
				int intNewRowNum = fnAppendEmptyRowToExcelSheet(pageObjects.strConfluenceSheetPath, "FailedTests")+1;
				fnInsertCellValueAtRowNum(pageObjects.strConfluenceSheetPath, "FailedTests", intNewRowNum, "Product_Name", pageObjects.strProductName, "");
				fnInsertCellValueAtRowNum(pageObjects.strConfluenceSheetPath, "FailedTests", intNewRowNum, "FeatureDetail", pageObjects.strFeatureName, "");
				fnInsertCellValueAtRowNum(pageObjects.strConfluenceSheetPath, "FailedTests", intNewRowNum, "TestCaseName", pageObjects.strGVScrDataPointer, pageObjects.strPDFPath);
				fnInsertCellValueAtRowNum(pageObjects.strConfluenceSheetPath, "FailedTests", intNewRowNum, "Remarks", strRemarks, "");
				fnInsertCellValueAtRowNum(pageObjects.strConfluenceSheetPath, "FailedTests", intNewRowNum, "Rel_Version", pageObjects.strReleaseVersion, "");
				fnInsertCellValueAtRowNum(pageObjects.strConfluenceSheetPath, "FailedTests", intNewRowNum, "Date_Time", fnGetCurrentDate("yyyy/mm/dd") + "; " + strCurrTime1, ""); }

		File strFile = new File(pageObjects.strHTMLPath + pageObjects.strHTMLName);
		if (strFile.exists() == true) {
			FileWriter strHTMLFile = new FileWriter(pageObjects.strHTMLPath + pageObjects.strHTMLName, true);
			strHTMLFile.write(strRows);
			strHTMLFile.close();  } 

		//Step to update the previous environments and user ids
		pageObjects.strPrevLoginUserID = pageObjects.strLoginUserID; 
		pageObjects.strAddUser = false; }
	
	
	
	//function to generate the HTML results row one by one
	public synchronized void fnGenerateHTMLRows_ORIGINAL(String strHTMLFilePath, String strTC, String strStatus, String strDealID, String strTime, String strRemarks) throws Throwable{
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
				
				pageObjects.strPDFPath =  strTempPath1 + strTC + ".PDF ";
				if (pageObjects.strAddUser == true) {
					File file1 = new File(pageObjects.strPDFPath);
					if (file1.exists()){
						String strTC1 = strTempPath1 + strTC + "_User_" + pageObjects.strLoginUserID + ".PDF";
						File file2 = new File(strTC1);
						if (file2.exists()==false) {
							if (file1.renameTo(file2)){
								System.out.println("[INFO]   File renamed with user id - " + strTC1); 
								strTC = strTC + "_User_" + pageObjects.strLoginUserID + ".PDF";
								pageObjects.strPDFPath =  strTC1; }
							else { System.out.println("[INFO]   Failed to rename the file - " + strTC); } } } }

			if (strStatus.trim().toLowerCase().contentEquals("pass") == true)
				strColor = "#006633"; // green color code
			else if (strStatus.trim().toLowerCase().contentEquals("fail") == true)
				strColor = "#D80000"; //red color code
	
			String strCurrTime1 = new SimpleDateFormat("HH.mm").format(new Date());
			String strRows = "<TR  style='background:#f2f2f2'" + ";text-align='left'>";
			
			strRows = strRows + "<TH BGCOLOR = '' ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=1> <a href=" + "\"" + pageObjects.strPDFPath + "\"" + "onclick = open(this.href, this.target, 'width=600, height=450, top=200, left=250'); return false; target=_blank><font color='" + strColor + "'>" + strTC + "</font></a> </FONT></TH>";
			strRows = strRows + "<TH BGCOLOR = '' ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=1><font color='" + strColor + "'>" + strStatus + "</font> </FONT></TH>";

			strRows = strRows + "<TH ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=1>" + strDealID + " </FONT></TH>";
			strRows = strRows + "<TH ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=1>" + strCurrTime1 + " </FONT></TH>";
			strRows = strRows + "<TH STYLE='WORD-WRAP:BREAK-WORD' ALIGN='LEFT'><FONT COLOR=BLACK FACE='Geneva, Verdana' SIZE=1>" + strRemarks + " </FONT></TH></TR>";
	
			//Steps to update the test failure details in ConfluenceDashBoardResults.xls - FailedTests sheet
			if (fnGetExcelData(pageObjects.strConfigFilepath, "DPMConfig", "UpdateConfluence", "ConfigValue").trim().toLowerCase().contentEquals("true"))
				if (strStatus.trim().toLowerCase().toLowerCase().contentEquals("fail")) {
					int intNewRowNum = fnAppendEmptyRowToExcelSheet(pageObjects.strConfluenceSheetPath, "FailedTests")+1;
					fnInsertCellValueAtRowNum(pageObjects.strConfluenceSheetPath, "FailedTests", intNewRowNum, "Product_Name", pageObjects.strProductName, "");
					fnInsertCellValueAtRowNum(pageObjects.strConfluenceSheetPath, "FailedTests", intNewRowNum, "FeatureDetail", pageObjects.strFeatureName, "");
					fnInsertCellValueAtRowNum(pageObjects.strConfluenceSheetPath, "FailedTests", intNewRowNum, "TestCaseName", pageObjects.strGVScrDataPointer, pageObjects.strPDFPath);
					fnInsertCellValueAtRowNum(pageObjects.strConfluenceSheetPath, "FailedTests", intNewRowNum, "Remarks", strRemarks, "");
					fnInsertCellValueAtRowNum(pageObjects.strConfluenceSheetPath, "FailedTests", intNewRowNum, "Rel_Version", pageObjects.strReleaseVersion, "");
					fnInsertCellValueAtRowNum(pageObjects.strConfluenceSheetPath, "FailedTests", intNewRowNum, "Date_Time", fnGetCurrentDate("yyyy/mm/dd") + "; " + strCurrTime1, ""); }
	
			File strFile = new File(pageObjects.strHTMLPath + pageObjects.strHTMLName);
			if (strFile.exists() == true) {
				FileWriter strHTMLFile = new FileWriter(pageObjects.strHTMLPath + pageObjects.strHTMLName, true);
				strHTMLFile.write(strRows);
				strHTMLFile.close();  } 
			
	//Step to update the previous environments and user ids
	pageObjects.strPrevLoginUserID = pageObjects.strLoginUserID; 
	pageObjects.strAddUser = false; }
	
	//Function to generate the EMAIL and trigger the mail to recipients====================================================================================================
	public void fnGeneateEmail(String stEnvironment) throws Throwable {
			BufferedReader br = null;
			final String emailSubject;
			stEnvironment = pageObjects.strEnvironment;
			String strTempPth = pageObjects.strHTMLPath + "\\\\" + pageObjects.strConfigFileName;

			pageObjects.intSheetRowNum = 0;
			if ((pageObjects.strTriggerEmail.trim().toLowerCase().contentEquals("yes")) && (pageObjects.intExecCount > 0)) {
				try {
					final String fromusername = fnGetExcelData(strTempPth, "eMailRecepient", stEnvironment, "FromReceipent");
					final String tousername = fnGetExcelData(strTempPth, "eMailRecepient", stEnvironment, "ToReceipent");
					final String ccusername = fnGetExcelData(strTempPth, "eMailRecepient", stEnvironment, "CcReceipent");
					final String authusername = fnGetExcelData(strTempPth, "eMailRecepient", stEnvironment, "AuthUserName");
					if(pageObjects.strPIVStatus == false)
						emailSubject = fnGetExcelData(strTempPth, "eMailRecepient", stEnvironment, "MailSubject");
					else
						emailSubject = "PIV - Automation Execution - Report";

					String host = fnGetExcelData(strTempPth, "eMailRecepient", stEnvironment, "MailHost");
					final String pwd = "";
					String mailText = "";
					System.out.println("[INFO]   from user - " + fromusername);
					System.out.println("[INFO]   to user - " + tousername);
					System.out.println("[INFO]   cc user - " + ccusername);
					System.out.println("[INFO]   auth user - " + authusername);
					System.out.println("[INFO]   host - " + host);
					
					Properties strOutLook = new Properties();
					strOutLook.put("mail.smtp.host",host);
					//strOutLook.put("mail.smtp.auth", "true");
					strOutLook.put("mail.smtp.port", "25");
					strOutLook.put("mail.debug", "true");
					strOutLook.put("mail.transport.protocol", "smtps");
					//strOutLook.put("mail.smtp.starttls.enable", "false");
					strOutLook.put("mail.smtp.starttls.enable", "true");

					Session mailsession = Session.getInstance(strOutLook, new javax.mail.Authenticator() {
					            protected PasswordAuthentication getPasswordAuthentication() {
					                return new PasswordAuthentication(authusername, pwd); } });
			
			        MimeMessage message = new MimeMessage(mailsession);
			        message.saveChanges();
			        message.setFrom(new InternetAddress(fromusername));
			        message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(tousername));
			        message.setRecipients(Message.RecipientType.CC, InternetAddress.parse(ccusername));
			        
			        if(pageObjects.strPIVStatus == false)
						br = new BufferedReader(new FileReader(pageObjects.strHTMLPath + pageObjects.strHTMLName));
					else if(pageObjects.strPIVStatus == true)
						br = new BufferedReader(new FileReader(pageObjects.PIVTestReportFile));

			        String sCurrentLine;
			        while ((sCurrentLine = br.readLine()) != null) {
			        	mailText = mailText + sCurrentLine; }
			        message.setSubject(emailSubject + " - " + fnGetCurrentDate("yyyy-mm-dd"));
			        message.setContent(mailText, "text/html");
			        Transport.send(message);
			        System.out.println("[INFO]   completed the e-mail function");
			        br.close(); }
				catch (Exception exp) {
					System.out.println("[ERROR]   *** *** *** *** Error encountered in triggering email - " + exp.getMessage()); } } 

	        //Step to delete the chrome driver & configuration excel file from execution folder
	        File file1 = new File(pageObjects.strHTMLPath+"\\\\"+"chromedriver.exe");
	        try {
	        if (file1.exists())
	        	file1.delete(); }
	        catch (Exception exp) {
	        	System.out.println("[ERROR]   Unable to delete the file - " + pageObjects.strHTMLPath+"\\\\"+"chromedriver.exe"); }
	        file1 = new File(pageObjects.strHTMLPath+"\\\\"+"peakxv_Configuration.xls");
	        try {
	        if (file1.exists())
	        	file1.delete(); }
	        catch (Exception exp) {
	        	System.out.println("[ERROR]   Unable to delete the file - " + pageObjects.strHTMLPath+"\\\\"+"peakxv_Configuration.xls"); }	
	        
	        System.setOut(stdout); }
	
	//function to generate the HTML report for test execution=============================================================================================================================
	public synchronized void fnGenerateHTMLFooterReport(String strHTMLFilePath) throws Throwable, Throwable {
			FileWriter strHTMLFile;
			String strVal="";
			String strFooter = "</TABLE> <p><br> <FONT COLOR=BLUE FACE='Geneva, Arial' SIZE=2> Note: Click on the Use_Case name link to open the screen shot <br><br> <FONT COLOR=BLACK FACE='Geneva, Arial' SIZE=2>Thanks, <br>PEAK XV - Automation. </br></FONT></p></html>"; 
			File strFile = new File(pageObjects.strHTMLPath + pageObjects.strHTMLName);
			if (strFile.exists() == true) {
				strFile.createNewFile();
				strHTMLFile = new FileWriter(pageObjects.strHTMLPath + pageObjects.strHTMLName, true);
				strHTMLFile.write(strFooter);
				strHTMLFile.close(); } 

			//Steps to generate the data for confluence update ==========================================================================================
			String strTempFilePath = strConfluenceSheetPath;
			int intRowCount  = fnGetExcelRowCount(pageObjects.strDataSheetPath, pageObjects.strInputSheetName);
			int intTotCases=0,intPass=0,intFail=0;
			for(int intTemp=0;intTemp<intRowCount;intTemp++){
				strVal = fnGetExcelCellData(pageObjects.strDataSheetPath, pageObjects.strInputSheetName, intTemp, "ExecutionFlag");
				String strVal1 = fnGetExcelCellData(pageObjects.strDataSheetPath, pageObjects.strInputSheetName, intTemp, "Test_Status");
				if (strVal.trim().length()>0) {
					if ((strVal.trim().toLowerCase().contentEquals("y")) || (strVal.trim().toLowerCase().contentEquals("executed"))) intTotCases = intTotCases + 1;
					if (strVal1.trim().toLowerCase().contentEquals("pass")) intPass = intPass + 1;
					if (strVal1.trim().toLowerCase().contentEquals("fail")) intFail = intFail + 1; } }

			strVal = fnGetExcelData(pageObjects.strHTMLPath + pageObjects.strConfigFileName, "DPMConfig", "strExecutionFlag", "ConfigValue");
			String strTxt = pageObjects.strFunctionality + ":" + strVal + ":" + intTotCases + ":" + intPass + ":" + intFail;
			fnWriteTxtFile(strConfluenceCurTxtPath + pageObjects.strFunctionality + ".txt", strTxt);
			System.out.println("[INFO]   Created the Confluence Update file for process - " + strConfluenceCurTxtPath + pageObjects.strFunctionality + ".txt");
	}
				
	//function to verify the value exists or not in the drop down/ select item====================================================================================================
	public int fnSelectValue(String strObject, String strVal) throws Throwable{
		int strFound = 0;
		Robot robot = new Robot();
		robot.delay(100);
		try{
			if (fnObjectExists(strObject, 1) != 0) {
				fnAppSync(1);
				wDriver.findElement(By.xpath(strObject)).click();
				new Select(wDriver.findElement(By.xpath(strObject))).selectByValue(strVal);
				fnAppSync(1);
				wDriver.findElement(By.xpath(strObject)).click();
				strFound = 1; } }
		catch(Exception exp){
			strFound = 0;
			pageObjects.strScnStatus = false;
			pageObjects.strGVErrMsg = "Failed to search for value - " + strVal + " from the drop down selection";
			pageObjects.strGVStepName = "Failed to select/find " + strVal + " value from the dropdown";
			fnScreenCapture(wDriver, strScreenPath.concat("peakxvDealBookFailure") + "-"+ pageObjects.strGVScrDataPointer);
			pageObjects.strScnStatus = true;
			System.out.println("[ERROR]    ****** Failed to select the value : " + strVal + ", from the drop down object ID - " + strObject); }
		return strFound; }
	
	//function to verify the value exists or not in the drop down/ select item====================================================================================================
	public int fnSelectValueByValue(String strObject, String strVal) throws Throwable{
		int strFound = 0;
		Robot robot = new Robot();
		robot.delay(100);
		try{
			if (fnObjectExists(strObject, 1) != 0) {
				new Select(wDriver.findElement(By.xpath(strObject))).selectByValue(strVal);
				System.out.println("[INFO]    Successfully selected the value : " + strVal + ", from Object : " + strObject);
				fnAppSync(1);
				strFound = 1; } }
		catch(Exception exp){
			strFound = 0;
			pageObjects.strScnStatus = false;
			pageObjects.strGVErrMsg = "Failed to select/find value - " + strVal + " from the drop down selection";
			pageObjects.strGVStepName = "Failed to select/find value - " + strVal + " from the drop down selection";
			fnScreenCapture(wDriver, strScreenPath.concat("peakxvDealBookFailure") + "-"+ pageObjects.strGVScrDataPointer);
			pageObjects.strScnStatus = true;
			System.out.println("[ERROR]    ****** Failed to select the value : " + strVal + ", from the drop down object ID - " + strObject); }
		return strFound; }
	
	//function to verify the value exists or not in the drop down/ select item====================================================================================================
	public int fnMultipleSelectByValue(String strObject, String strVal) throws Throwable{
		fnAppSync(1);
		int strFound = 0;
		try{
			if (fnObjectExists(strObject, 1) != 0) {
				fnAppSync(1);
				//wDriver.findElement(By.xpath(strObject)).click();
				
				new Select(wDriver.findElement(By.xpath(strObject))).selectByVisibleText(strVal);
				
				//wDriver.findElement(By.xpath(strObject)).click();
				strFound = 1; } }
		catch(Exception exp){
			strFound = 0;
			pageObjects.strScnStatus = false;
			pageObjects.strGVErrMsg = "Failed to search for value - " + strVal + "in the drop down selection";
			pageObjects.strGVStepName = "Failed to select/find " + strVal + " value from the dropdown";
			fnScreenCapture(wDriver, strScreenPath.concat("peakxvDealBookFailure") + "-"+ pageObjects.strGVScrDataPointer);
			pageObjects.strScnStatus = true;
			System.out.println("[ERROR]    ****** Failed to select the value : " + strVal + ", from the drop down object ID - " + strObject); }
		return strFound; }
	
	//Function to find the given string value is a numeric or not=============================================================================================================================
		public boolean isNumeric(String strVal) {
			
			return strVal.matches("^(?:(?:\\-{1})?\\d+(?:\\.{1}\\d+)?)$");
			
			//return NumberUtils.isNumber(strVal);
			
		    //try { Long.parseLong(strVal); }
		    //catch(NumberFormatException nfe) { return false; }
		    //return true; 
		    }

	//Function to find the given list item exists in the dropdown object or not
	public String fnGetListElementText(String strObject, int strVal) throws Throwable{
		Select dropdown = new Select(wDriver.findElement(By.xpath(strObject)));
		if (fnObjectExists(strObject, 1) != 0) {
			//Get all options
			List<WebElement> strDrpDwn = dropdown.getOptions();
			if (strVal <= strDrpDwn.size())
				return strDrpDwn.get(strVal).getText();
			else
				return "0"; }
		else
			return "0"; }
	
	//Function to find the find the given string is a exponential value or not ===============================================================================================================
	public boolean isScientificNotation(String numberString) {
	    // Validate number
	    try {
	        new BigDecimal(numberString);
	    } 
	    catch (NumberFormatException e) {
	        return false; }
	    // Check for scientific notation
	    return numberString.toUpperCase().contains("E"); }

	
	//Function to copy the file from source to traget folder
	@SuppressWarnings("resource")
	public synchronized void fnCopyFile(String strSource, String strTarget) throws Throwable {

		String strTempPath = strSource;
		String strTempPath1 = "C";
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
		strSource = strTempPath1;
		
		strTempPath = strTarget;
		strTempPath1 = "\\\\";
		chr1 = (char) 92;
		intTemp=0;
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
		strTarget = strTempPath1 + pageObjects.strDataSheetPath;
				
		File srcFile = new File(strSource);
		File trgFile = new File(strTarget);
		
		if (!trgFile.exists())
			trgFile.createNewFile();
		fnAppSync(1);
		
		FileChannel source = null;
		FileChannel target = null;

		source = new FileInputStream(srcFile).getChannel();
		target = new FileOutputStream(trgFile).getChannel();

		target.transferFrom(source,  0, source.size());
		System.out.println("[INFO]   Copied the source file - " + strSource + ", to traget - " + strTarget);
		source.close();
		target.close(); }

	//Function get the values from excel sheet and compare against the application - string objects
	public void fnStringObjCompareValues(String strDataSheetPath,String dataSheet,String dataPointer, String strFieldName, String strColname, String strObject,String strAttribute) throws Throwable {
		String strActVal = null, strExpVal="";

		//Step to read the values from DB & application ui, compare the results
		if (pageObjects.strDBCompareVal == true){
			String strDBTablName="", strDbColName="", strDBCondColName="", strDBQuery="", strDBRepQuery="";
			strDBTablName = fnGetExcelData(strDataSheetPath, "CDS_DefQuery", strColname, "DB_TableName");
			strDbColName = fnGetExcelData(strDataSheetPath, "CDS_DefQuery", strColname, "DB_ColName");
			strDBCondColName = fnGetExcelData(strDataSheetPath, "CDS_DefQuery", strColname, "DB_Cond_ColName");
			strDBQuery = fnGetExcelData(strDataSheetPath, "CDS_DefQuery", strColname, "DB_Query");
			strDBRepQuery = strDBQuery;
			
			strDBRepQuery = strDBRepQuery.replace("DB_ColName", strDbColName);
			strDBRepQuery = strDBRepQuery.replace("DB_TableName", strDBTablName);
			strDBRepQuery = strDBRepQuery.replace("DB_Cond_ColName", strDBCondColName);
			strDBRepQuery = strDBRepQuery.replace("DB_VALUE", "ASSICURAZIONI GEN (SUB)");

			System.out.println("[INFO]    DB_Tabl Name : " + strDBTablName);
			System.out.println("[INFO]    DB_Col Name : " + strDbColName);
			System.out.println("[INFO]    DB_Cond_Col Name : " + strDBCondColName);
			System.out.println("[INFO]    DB_Query : " + strDBQuery);
			System.out.println("[INFO]    DB_Rep_Query : " + strDBRepQuery);
			System.out.println("[INFO]    DB - Env. " + pageObjects.strLoginEnv);
			
			strExpVal = fnReadDBValues(strDBRepQuery, pageObjects.strLoginEnv, 1);
			System.out.println("[INFO]    DB_Rep_Query - Result  : " + strExpVal);
			
			if (fnObjectExists(strObject, 2)!=0)
				if (strExpVal.isEmpty() == false) {
					//Step to get the text from the SPAN object type
					String strNodeVal = fnGetObjectAttributeValue(strObject, "nodeName");
					if (strNodeVal.isEmpty() == false) {
						if ((strNodeVal.trim().toLowerCase().contentEquals("span") == true) || (strNodeVal.trim().toLowerCase().contentEquals("button") == true))
							strAttribute = "outerText";
						
						//step to read the object attribute property value
						strActVal = fnGetObjectAttributeValue(strObject, strAttribute);
						if (strExpVal.trim().toLowerCase().contains(strActVal.trim().toLowerCase()) || strActVal.trim().toLowerCase().contains(strExpVal.trim().toLowerCase()))
							System.out.println("[INFO]    == == == String DB Comparison PASSED - " + strFieldName + " - Actual Value : " + strActVal + " - DB Expected Value : " + strExpVal);
						else {
							System.out.println("[ERROR]    ****** String DB Comparison FAILED - " + strFieldName + " - Actual Value : " + strActVal + " - DB Expected Value : " + strExpVal);
							pageObjects.strCompareVal = false;
							if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
							pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + " Field name : " + strFieldName + " - Actual Value : " + strActVal + " - DB Expected Value : " + strExpVal + "; \n "; } } 
					else {
						System.out.println("[ERROR]    ****** Failed to retrieve the value from field : " + strFieldName);
						if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
						pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + "Failed to retrieve the value from String field : " + strFieldName + "; \n "; } }
		} else
		if (pageObjects.strDBCompareVal == false){
			//step to retrieve the trade data from excel and book detail screen and compare
			strExpVal = fnGetExcelData(strDataSheetPath, dataSheet, dataPointer, strColname).trim().toLowerCase();
			
			if (fnObjectExists(strObject, 2)!=0)
			if (strExpVal.isEmpty() == false) {
				//Step to get the text from the SPAN object type
				String strNodeVal = fnGetObjectAttributeValue(strObject, "nodeName").trim();
				if (strNodeVal.isEmpty() == false) {
					if ((strNodeVal.trim().toLowerCase().contentEquals("span") == true) || (strNodeVal.trim().toLowerCase().contentEquals("button") == true))
						strAttribute = "outerText";
					
					//step to read the object attribute property value
					strActVal = fnGetObjectAttributeValue(strObject, strAttribute).trim();
					if (strExpVal.trim().toLowerCase().contains(strActVal.trim().toLowerCase()) || strActVal.trim().toLowerCase().contains(strExpVal.trim().toLowerCase()))
						System.out.println("[INFO]    == == == String Comparison PASSED - " + strFieldName + " - Actual Value : " + strActVal + " - Expected Value : " + strExpVal);
					else {
						System.out.println("[ERROR]    ****** String Comparison FAILED - " + strFieldName + " - Actual Value : " + strActVal + " - Expected Value : " + strExpVal);
						pageObjects.strCompareVal = false;
						if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
						pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + " Field name : " + strFieldName + " - Actual Value : " + strActVal + " - Expected Value : " + strExpVal + "; \n "; } } 
				else {
					System.out.println("[ERROR]    ****** Failed to retrieve the value from field : " + strFieldName);
					if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
					pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + "Failed to retrieve the value from String field : " + strFieldName + "; \n "; } } } }

	//Function get the values from excel sheet and compare against the application - DATE type objects
	public void fnDateObjCompareValues(String strDataSheetPath,String dataSheet,String dataPointer, String strFieldName, String strColname, String strObject,String strAttribute, String strDateFormat) throws Throwable {
		String strActVal = null, strExpVal="";
		//Step to read the values from DB & application ui, compare the results
		if ((pageObjects.strDBCompareVal == true) && (pageObjects.strDBCompSheetName.trim().isEmpty() == false)){
			String strDBTablName="", strDbColName="", strDBCondColName="", strDBQuery="", strDBRepQuery="";
			strDBTablName = fnGetExcelData(strDataSheetPath, "CDS_DefQuery", strColname, "DB_TableName");
			strDbColName = fnGetExcelData(strDataSheetPath, "CDS_DefQuery", strColname, "DB_ColName");
			strDBCondColName = fnGetExcelData(strDataSheetPath, "CDS_DefQuery", strColname, "DB_Cond_ColName");
			strDBQuery = fnGetExcelData(strDataSheetPath, "CDS_DefQuery", strColname, "DB_Query");
			strDBRepQuery = strDBQuery;
			
			strDBRepQuery = strDBRepQuery.replace("DB_ColName", strDbColName);
			strDBRepQuery = strDBRepQuery.replace("DB_TableName", strDBTablName);
			strDBRepQuery = strDBRepQuery.replace("DB_Cond_ColName", strDBCondColName);
			strDBRepQuery = strDBRepQuery.replace("DB_VALUE", "ASSICURAZIONI GEN (SUB)");

			System.out.println("[INFO]    DB_Tabl Name : " + strDBTablName);
			System.out.println("[INFO]    DB_Col Name : " + strDbColName);
			System.out.println("[INFO]    DB_Cond_Col Name : " + strDBCondColName);
			System.out.println("[INFO]    DB_Query : " + strDBQuery);
			System.out.println("[INFO]    DB_Rep_Query : " + strDBRepQuery);
			System.out.println("[INFO]    DB - Env. " + pageObjects.strLoginEnv);
			
			strExpVal = fnReadDBValues(strDBRepQuery, pageObjects.strLoginEnv, 1);
			System.out.println("[INFO]    DB_Rep_Query - Result  : " + strExpVal);
			
			if (fnObjectExists(strObject, 2)!=0)
				if (strExpVal.toLowerCase().toLowerCase().contentEquals("null") == false)
				if ((strExpVal.isEmpty() == false) || (pageObjects.strDefaultValues == true)) {
					//step to read the object attribute property value
					strActVal = fnGetObjectAttributeValue(strObject, strAttribute);
					if (strColname.trim().toLowerCase().contentEquals("swap_expirydate")) {
						if (strActVal.trim().toLowerCase().contains("|")) {
							String[] strVal2 = strActVal.split("\\|");
							strActVal = strVal2[1].trim(); } }

					if (pageObjects.strDefaultValues == true){
						if (strActVal.trim().length()<1) strActVal = "null";
						if (strExpVal.trim().length()<1) strExpVal = "null"; }
					
					if ((strActVal != null) || (pageObjects.strDefaultValues == true)) {
						if (strExpVal.trim().toLowerCase().contentEquals(strActVal.trim().toLowerCase()) == true)
							System.out.println("[INFO]    == == == Date DB Comparison PASSED - " + strFieldName + " - Actual Value : " + strActVal + " - DB Expected Value : " + strExpVal);
						else {
							System.out.println("[ERROR]    ****** Date DB Comparison FAILED - " + strFieldName + " - Actual Value : " + strActVal + " - DB Expected Value : " + strExpVal);
							pageObjects.strCompareVal = false;
							if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
								pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + " Field name : " + strFieldName + " - Actual Value : " + strActVal + " - DB Expected Value : " + strExpVal + "; \n "; } }
					else {
						System.out.println("[ERROR]    ****** Failed to retrieve the value from Date field : " + strFieldName);
						if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
						pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + "Failed to retrieve the value from Date field : " + strFieldName + "; \n "; } }
		} else
		if (pageObjects.strDBCompareVal == false) {
			//step to retrieve the trade date from excel and book detail screen and compare
			strExpVal = fnGetExcelData(strDataSheetPath, dataSheet, dataPointer, strColname);
					
			if (fnObjectExists(strObject, 2)!=0)
			if (strExpVal.toLowerCase().toLowerCase().contentEquals("null") == false)
			if ((strExpVal.isEmpty() == false) || (pageObjects.strDefaultValues == true)) {
				//step to read the object attribute property value
				strActVal = fnGetObjectAttributeValue(strObject, strAttribute);
				if (strActVal.trim().toLowerCase().contains("|")) {
					String[] strVal2 = strActVal.split("\\|");
					strActVal = strVal2[1].trim(); }
				
				if (pageObjects.strDefaultValues == true){
					if (strActVal.trim().length()<1) strActVal = "null";
					if (strExpVal.trim().length()<1) strExpVal = "null"; }
				
				if ((strActVal != null) || (pageObjects.strDefaultValues == true)) {
					if (strActVal.trim().substring(0, 1).contentEquals("0")) strActVal = strActVal.trim().substring(1, strActVal.length());
					if (strExpVal.trim().substring(0, 1).contentEquals("0")) strExpVal = strExpVal.trim().substring(1, strExpVal.length());
					
					if (strExpVal.trim().toLowerCase().contentEquals(strActVal.trim().toLowerCase()) == true)
						System.out.println("[INFO]    == == == Date Comparison PASSED - " + strFieldName + " - Actual Value : " + strActVal + " - Expected Value : " + strExpVal);
					else {
						System.out.println("[ERROR]    ****** Date Comparison FAILED - " + strFieldName + " - Actual Value : " + strActVal + " - Expected Value : " + strExpVal);
						pageObjects.strCompareVal = false;
						if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
							pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + " Field name : " + strFieldName 
									+ " - Actual Value : " + strActVal + " - Expected Value : " + strExpVal + "; \n "; } }
				else {
					System.out.println("[ERROR]    ****** Failed to retrieve the value from Date field : " + strFieldName);
					if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
					pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + "Failed to retrieve the value from Date field : " + strFieldName + "; \n "; } } } }

	//Function get the values from excel sheet and compare against the application - string objects
	public void fnListObjCompareValues(String strDataSheetPath,String dataSheet,String dataPointer, String strFieldName, String strColname, String strObject, String strAttribute) throws Throwable {
		
		String strActVal = null, strExpVal="";

		//Step to read the values from DB & application ui, compare the results
		if ((pageObjects.strDBCompareVal == true) && (pageObjects.strDBCompSheetName.trim().isEmpty() == false)){
			String strDBTablName="", strDbColName="", strDBCondColName="", strDBQuery="", strDBRepQuery="";
			strDBTablName = fnGetExcelData(strDataSheetPath, pageObjects.strDBCompSheetName, strColname, "DB_TableName");
			strDbColName = fnGetExcelData(strDataSheetPath, pageObjects.strDBCompSheetName, strColname, "DB_ColName");
			strDBCondColName = fnGetExcelData(strDataSheetPath, pageObjects.strDBCompSheetName, strColname, "DB_Cond_ColName");
			strDBQuery = fnGetExcelData(strDataSheetPath, pageObjects.strDBCompSheetName, strColname, "DB_Query");
			strDBRepQuery = strDBQuery;
			
			strDBRepQuery = strDBRepQuery.replace("DB_ColName", strDbColName);
			strDBRepQuery = strDBRepQuery.replace("DB_TableName", strDBTablName);
			strDBRepQuery = strDBRepQuery.replace("DB_Cond_ColName", strDBCondColName);
			strDBRepQuery = strDBRepQuery.replace("DB_VALUE", "ASSICURAZIONI GEN (SUB)");

			System.out.println("[INFO]    DB_Tabl Name : " + strDBTablName);
			System.out.println("[INFO]    DB_Col Name : " + strDbColName);
			System.out.println("[INFO]    DB_Cond_Col Name : " + strDBCondColName);
			System.out.println("[INFO]    DB_Query : " + strDBQuery);
			System.out.println("[INFO]    DB_Rep_Query : " + strDBRepQuery);
			System.out.println("[INFO]    DB - Env. " + pageObjects.strLoginEnv);
			
			strExpVal = fnReadDBValues(strDBRepQuery, pageObjects.strLoginEnv, 1);
			System.out.println("[INFO]    DB_Rep_Query - Result  : " + strExpVal);
			
			String strConfigVal = "";
			//strConfigVal = fnGetExcelData(pageObjects.strConfigFilepath, "ListValues", strExpVal, "ConfigValue");
			String strExpVal1 = strExpVal;
			if (strConfigVal.trim().length() > 0) strExpVal = strConfigVal.trim().toLowerCase();
			
			if (fnObjectExists(strObject, 1) != 0)
			if (strExpVal.isEmpty() == false) {
				if (fnGetObjectAttributeValue(strObject, "nodeName").trim().toLowerCase().contentEquals("select") == true) {
				//step to read the object attribute property value
				strActVal = fnGetObjectAttributeValue(strObject, strAttribute);
				int strFound = 0;
				if (strExpVal.trim().toLowerCase().contentEquals(strActVal.trim().toLowerCase()) || strActVal.trim().toLowerCase().contentEquals(strExpVal.trim().toLowerCase()))
					strFound = 1;
				else if (strExpVal1.trim().toLowerCase().contentEquals(strActVal.trim().toLowerCase()) || strActVal.trim().toLowerCase().contentEquals(strExpVal1.trim().toLowerCase()))
					strFound = 1;
	
				if (strFound == 1)
					System.out.println("[INFO]    == == == List String DB Comparison PASSED - " + strFieldName + " - Actual Value : " + strActVal + " - DB Expected Value : " + strExpVal);
				else {
					System.out.println("[ERROR]    ****** List String DB Comparison FAILED - " + strFieldName + " - Actual Value : " + strActVal + " - DB Expected Value : " + strExpVal);
					pageObjects.strCompareVal = false;
					if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
						pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + " Field name : " + strFieldName + " - Actual Value : " + strActVal + " - DB Expected Value : " + strExpVal + "; \n "; } } 
			else {
				System.out.println("[ERROR]    ****** Failed to retrieve the value from field : " + strFieldName);
				if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
				pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + "Failed to retrieve the value from List field : " + strFieldName + "; \n "; } }
		} else
		if (pageObjects.strDBCompareVal == false){

			//step to retrieve the trade date from excel and book detail screen and compare
			strExpVal = fnGetExcelData(strDataSheetPath, dataSheet, dataPointer, strColname).trim();
			String strConfigVal = "";
			//strConfigVal = fnGetExcelData(pageObjects.strConfigFilepath, "ListValues", strExpVal, "ConfigValue");
			String strExpVal1 = strExpVal;
			if (strConfigVal.trim().length() > 0) strExpVal = strConfigVal.trim().toLowerCase();
			
			if (fnObjectExists(strObject, 1) != 0)
			if (strExpVal.isEmpty() == false) {
				if (fnGetObjectAttributeValue(strObject, "nodeName").trim().toLowerCase().contentEquals("select") == true) {
				//step to read the object attribute property value
				strActVal = fnGetObjectAttributeValue(strObject, strAttribute).trim();
				int strFound = 0;
				if (strExpVal.trim().toLowerCase().contentEquals(strActVal.trim().toLowerCase()) || strActVal.trim().toLowerCase().contentEquals(strExpVal.trim().toLowerCase()))
					strFound = 1;
				else if (strExpVal1.trim().toLowerCase().contentEquals(strActVal.trim().toLowerCase()) || strActVal.trim().toLowerCase().contentEquals(strExpVal1.trim().toLowerCase()))
					strFound = 1;
	
				if (strFound == 1)
					System.out.println("[INFO]    == == == List String Comparison PASSED - " + strFieldName + " - Actual Value : " + strActVal + " - Expected Value : " + strExpVal);
				else {
					System.out.println("[ERROR]    ****** List String Comparison FAILED - " + strFieldName + " - Actual Value : " + strActVal + " - Expected Value : " + strExpVal);
					pageObjects.strCompareVal = false;
					if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
						pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + " Field name : " + strFieldName + " - Actual Value : " + strActVal + " - Expected Value : " + strExpVal + "; \n "; } } 
			else {
				System.out.println("[ERROR]    ****** Failed to retrieve the value from field : " + strFieldName);
				if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
				pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + "Failed to retrieve the value from List field : " + strFieldName + "; \n "; } } } }

	//Function get the values from excel sheet and compare against the application - string objects
	public void fnNumericObjCompareValues(String strDataSheetPath,String dataSheet,String dataPointer, String strFieldName, String strColname, String strObject, String strAttribute) throws Throwable {
		
		String strActVal = null, strExpVal="";
		//Step to read the values from DB & application ui, compare the results
		
		if ((pageObjects.strDBCompareVal == true) && (pageObjects.strDBCompSheetName.trim().isEmpty() == false)){
			String strDBTablName="", strDbColName="", strDBCondColName="", strDBQuery="", strDBRepQuery="";
			strDBTablName = fnGetExcelData(strDataSheetPath, "CDS_DefQuery", strColname, "DB_TableName");
			strDbColName = fnGetExcelData(strDataSheetPath, "CDS_DefQuery", strColname, "DB_ColName");
			strDBCondColName = fnGetExcelData(strDataSheetPath, "CDS_DefQuery", strColname, "DB_Cond_ColName");
			strDBQuery = fnGetExcelData(strDataSheetPath, "CDS_DefQuery", strColname, "DB_Query");
			strDBRepQuery = strDBQuery;
			
			strDBRepQuery = strDBRepQuery.replace("DB_ColName", strDbColName);
			strDBRepQuery = strDBRepQuery.replace("DB_TableName", strDBTablName);
			strDBRepQuery = strDBRepQuery.replace("DB_Cond_ColName", strDBCondColName);
			strDBRepQuery = strDBRepQuery.replace("DB_VALUE", "ASSICURAZIONI GEN (SUB)");

			System.out.println("[INFO]    DB_Tabl Name : " + strDBTablName);
			System.out.println("[INFO]    DB_Col Name : " + strDbColName);
			System.out.println("[INFO]    DB_Cond_Col Name : " + strDBCondColName);
			System.out.println("[INFO]    DB_Query : " + strDBQuery);
			System.out.println("[INFO]    DB_Rep_Query : " + strDBRepQuery);
			System.out.println("[INFO]    DB - Env. " + pageObjects.strLoginEnv);
			
			strExpVal = fnReadDBValues(strDBRepQuery, pageObjects.strLoginEnv, 1);
			System.out.println("[INFO]    DB_Rep_Query - Result  : " + strExpVal);

			DecimalFormat myFormatter = new DecimalFormat("#############.#####");
			if (fnObjectExists(strObject, 2)!=0)
			if (strExpVal.trim().isEmpty() == false) {
				//step to read the object attribute property value
				strActVal = fnGetObjectAttributeValue(strObject, strAttribute);
				strActVal = strActVal.replaceAll("$", "").trim();
				if (strActVal.trim().isEmpty() == false) {
					String strTemp, strTemp1;
					strTemp = myFormatter.format(Double.parseDouble(strActVal.replaceAll(",", "").trim()));
					strTemp1 = myFormatter.format(Double.parseDouble(strExpVal.replaceAll(",", "").trim()));
	
					if (strTemp.equals(strTemp1) == true)
						System.out.println("[INFO]    == == == Numeric DB Comparison PASSED - " + strFieldName + " - Actual Value : " + strActVal + " - DB Expected Value : " + strExpVal);
					else {
						System.out.println("[ERROR]    ****** Numeric DB Comparison FAILED - " + strFieldName + " - Actual Value : " + strActVal + " - DB Expected Value : " + strExpVal);
						pageObjects.strCompareVal = false;
						if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
							pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + " Field name : " + strFieldName + " - Actual Value : " + strActVal + " - DB Expected Value : " + strExpVal + "; \n "; } }
				else
					System.out.println("[ERROR]    ****** Failed to read the value from the Field : " + strFieldName + ", Object Locator ID : " + strObject); } }
		else
		if (pageObjects.strDBCompareVal == false){
			
			//step to retrieve the trade date from excel and book detail screen and compare ========================================================================
			strExpVal = fnGetExcelData(strDataSheetPath, dataSheet, dataPointer, strColname).trim().replace("$", "");
			strExpVal = strExpVal.replace("\"", "").trim();
			strExpVal = strExpVal.replace(",", "").replace(" ", "");
			
			DecimalFormat myFormatter = new DecimalFormat("#############.#####");
			if (fnObjectExists(strObject, 2)!=0)
			if (strExpVal.trim().isEmpty() == false) {
				//step to read the object attribute property value ================================================================================================
				strActVal = fnGetObjectAttributeValue(strObject, strAttribute).trim();
				strActVal = strActVal.replace("$", "").trim();
				strActVal = strActVal.replaceAll("\"", "");
				strActVal = strActVal.replaceAll(",", "").trim();

				System.out.println("[INFO]    strActVal : " + strActVal + " ====== strExpVal : " + strExpVal);
				if (strActVal.trim().isEmpty() == false) {
					String strTemp, strTemp1;
					strTemp = myFormatter.format(Double.parseDouble(strActVal.replaceAll(",", "").trim()));
					strTemp1 = myFormatter.format(Double.parseDouble(strExpVal.replaceAll(",", "").trim()));
	
					if (strTemp.equals(strTemp1) == true)
						System.out.println("[INFO]    == == == Numeric Comparison PASSED - " + strFieldName + " - Actual Value : " + strActVal + " - Expected Value : " + strExpVal);
					else {
						System.out.println("[ERROR]    ****** Numeric Comparison FAILED - " + strFieldName + " - Actual Value : " + strActVal + " - Expected Value : " + strExpVal);
						pageObjects.strCompareVal = false;
						if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
							pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + " Field name : " + strFieldName + " - Actual Value : " + strActVal + " - Expected Value : " + strExpVal + "; \n "; } }
				else
					System.out.println("[ERROR]    ****** Failed to read the value from the Field : " + strFieldName + ", Object Locator ID : " + strObject); } } }
	
	//Function get the values from excel sheet and compare against the application - string objects
	public void fnCheckObjCompareValues(String strDataSheetPath,String dataSheet,String dataPointer, String strFieldName, String strColname, String strObject, String strAttribute) throws Throwable {
		//step to retrieve the trade date from excel and book detail screen and compare
		String strExpVal = fnGetExcelData(strDataSheetPath, dataSheet, dataPointer, strColname);
		String strActVal = null;
		if (strExpVal.trim().isEmpty() == false) {
			//step to read the object attribute property value
			strActVal = fnGetObjectAttributeValue(strObject, strAttribute);
			if (strActVal.trim().isEmpty() == false) {
				if ((strActVal.trim().toLowerCase().equals(strExpVal) == true) || (strActVal.trim().toLowerCase().equals("on") == true))
					System.out.println("[INFO]    == == == Check Option Comparison PASSED - " + strFieldName + " - Actual Value : " + strActVal + " - Expected Value : " + strExpVal);
				else {
					System.out.println("[ERROR]    ****** Check Option Comparison FAILED - " + strFieldName + " - Actual Value : " + strActVal + " - Expected Value : " + strExpVal);
					pageObjects.strCompareVal = false;
					if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
						pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + " Field name : " + strFieldName + " - Actual Value : " + strActVal + " - Expected Value : " + strExpVal + "; \n "; } }
			else
				System.out.println("[ERROR]    ****** Failed to read the value from the Field : " + strFieldName + ", Object Locator ID : " + strObject); } }

	//Function enter the values for date field - Date Objects
	public void fnDateObjEnterValue(String strObject, String strVal, int strWaitTime) throws Throwable{
		if (fnObjectExists(objTradeDate, 1) != 0) {
			if (strVal.trim().toLowerCase().contentEquals("today"))
				strVal = fnGetCurrentDate("yyyy-mm-dd");
			wDriver.findElement(By.xpath(objTradeDate)).clear();
			if (strWaitTime > 0) fnAppSync(strWaitTime);
			wDriver.findElement(By.xpath(objTradeDate)).sendKeys(strVal);
			//wDriver.findElement(By.xpath(objStartDate)).sendKeys(Keys.ENTER);
			wDriver.findElement(By.xpath(objStartDate)).sendKeys(Keys.TAB); } }

	//Function to search for the element from the drop down - MARKIT WIRE - EXCLUSIVE
	public int fnMrktSearchList(String strObj, String strVal) throws Throwable {
		fnAppSync(1);
		List<WebElement> eleList  = wDriver.findElements(By.xpath(strObj));
		int intListCnt  = eleList.size();
		String samp;
		int intFound = -1;
		for (int intTemp1 = 0; intTemp1 < intListCnt; intTemp1++) {
			samp = eleList.get(intTemp1).getAttribute("outerText");
			//System.out.println("list of samp ... " + intTemp1 + " - " + samp);
			if (samp.trim().toLowerCase().contentEquals(strVal.trim().toLowerCase())){
				if (samp.trim().length() == strVal.trim().length())
				{	System.out.println("[INFO]    FOUND ***** list element ... " + intTemp1 + " - " + samp);
					intFound = intTemp1; } }
			if (intFound != -1)
				break; }
		
		if (intFound != -1) {	
			Robot robot = new Robot();
			boolean strVisStatus;
			for (int intLoop = 0; intLoop < 3; intLoop++) {
				strVisStatus = eleList.get(intFound).isDisplayed();
				//System.out.println("PAGE UP - Visible .... " + strVisStatus);
				if (strVisStatus == false) {
					robot.keyPress(KeyEvent.VK_PAGE_UP);
					robot.keyRelease(KeyEvent.VK_PAGE_UP);
					//fnAppSync(1); 
					}
				else intLoop = 100; }
			for (int intLoop = 0; intLoop < 10; intLoop++) {
				strVisStatus = eleList.get(intFound).isDisplayed();
				//System.out.println("PAGE DOWN - Visible .... " + strVisStatus);
				if (strVisStatus == false) {
					robot.keyPress(KeyEvent.VK_PAGE_DOWN);
					robot.keyRelease(KeyEvent.VK_PAGE_DOWN);
					fnAppSync(1);
					}
				else intLoop = 100; }
			fnAppSync(2);
			eleList.get(intFound).click(); }
		else { System.out.println("[ERROR]    NOT FOUND ***** list element ... " + strVal); }
			//fnAppSync(1);
			return intFound; }
	
	//FUNCTION TO GET THE CURRENT DATE in given format
	public String fnGetCurrentDate1(String strDateFormat){
		DateFormat df = new SimpleDateFormat(strDateFormat);
       Date dateobj = new Date();
       String strdate = df.format(dateobj);
       return strdate; }
	
	//FUNCTION TO Update the ConfluenceDashboardResults.xls file into Confluence page
	public void fnUpdateConfluence(String strFilePath, String strDataSheet) throws Throwable {
		external_actions.loginLogout fnLoginModule = new external_actions.loginLogout();
		System.out.println("[INFO]    Reading UpdateConfigVal from Config sheet.....");
		String strUpdateConfl = String.valueOf(fnGetExcelData(pageObjects.strConfigFilepath, "DPMConfig", "UpdateConfluence", "ConfigValue"));
		System.out.println("[INFO]    Update Confluence sheet flag .. " + strUpdateConfl);
		if (strUpdateConfl.trim().toLowerCase().contentEquals("true")) {
				
			//Steps to take the backup of the confluencedashboardresults.xls file into the results folder
			System.out.println("[INFO]    Copying the initial source to the results folder .....");
			FileUtils.copyFile(new File(pageObjects.strConfluenceSheetPath), new File(pageObjects.strHTMLPath + "ConfluenceDashBoardResults.xls"));

			//Step to get the tests pass/ fail count
			int intRwCnt = fnGetExcelRowCount(strFilePath, strDataSheet);
			double intPassCnt=0, intFailCnt=0;

			System.out.println("[INFO]    Counting Pass and Failed count from the execution results sheet .....");
			for(int intLoop=0;intLoop<=intRwCnt;intLoop++) {
				String strStatus = fnGetExcelCellData(strFilePath, strDataSheet, intLoop, "Test_Status");
				String strExecFlag = fnGetExcelCellData(strFilePath, strDataSheet, intLoop, "ExecutionFlag");
				//System.out.println("Row No. " + intLoop + ", Status - " + strStatus + ", ExecutionFlag - " + strExecFlag);
				if (strExecFlag.trim().toLowerCase().contentEquals("executed"))
					if (strStatus.trim().toLowerCase().contentEquals("pass"))
						intPassCnt++;
					else if (strStatus.trim().toLowerCase().contentEquals("fail"))
						intFailCnt++; }

			//System.out.println("Total Act. Pass Cnt. " + intPassCnt);
			//System.out.println("Total Act. Fail Cnt. " + intFailCnt);
			String strProductName = pageObjects.strProductName;
			if (strDataSheet.trim().toLowerCase().contentEquals("testcase_feestab"))
				pageObjects.strProductName = "Fees_Tab";
			else if (strDataSheet.trim().toLowerCase().contentEquals("eqo_holidaypaymentstab"))
				pageObjects.strProductName = "HolidayPayments_Tab";
			else if ((strDataSheet.trim().toLowerCase().contentEquals("testcase_notepadtab_amend")) || (strDataSheet.trim().toLowerCase().contentEquals("testcase_notepadtab_booking"))) 
				pageObjects.strProductName = "Notepad_Tab";
			else if (strDataSheet.trim().toLowerCase().contentEquals("testcase_notepadtab_reverse")) 
				pageObjects.strProductName = "Notepad_Tab";
			else if (strDataSheet.trim().toLowerCase().contentEquals("testcase_optionalterminationtab")) 
				pageObjects.strProductName = "OptionalTermination_Tab";

			pageObjects.intSheetRowNum = 1;
			String strExeCnt="", strPassCnt="", strFailCnt="";
			System.out.println("[INFO]    Reading executed cell value for product - " + pageObjects.strProductName);
			strExeCnt = fnGetExcelData(pageObjects.strConfluenceSheetPath, "ConfluenceResults", pageObjects.strProductName, "Executed");
			if (strExeCnt.trim().length() < 1) strExeCnt = "0";
			System.out.println("[INFO]    Reading passed cell value for product - " + pageObjects.strProductName);
			strPassCnt = fnGetExcelData(pageObjects.strConfluenceSheetPath, "ConfluenceResults", pageObjects.strProductName, "Passed");
			if (strPassCnt.trim().length() < 1) strPassCnt = "0";
			System.out.println("[INFO]    Reading failed cell value for product - " + pageObjects.strProductName);
			strFailCnt = fnGetExcelData(pageObjects.strConfluenceSheetPath, "ConfluenceResults", pageObjects.strProductName, "Failed");
			if (strFailCnt.trim().length() < 1) strFailCnt = "0";

			DecimalFormat df2 = new DecimalFormat(".##");
			//DecimalFormat df1 = new DecimalFormat("#");
			double intTotCnt = Double.parseDouble(strExeCnt) + intPassCnt + intFailCnt;
			intPassCnt = intPassCnt + Double.parseDouble(strPassCnt);
			intFailCnt = intFailCnt + Double.parseDouble(strFailCnt);
			Double intPassPer = intPassCnt / intTotCnt;
			Double intFailPer = intFailCnt / intTotCnt;
			
			System.out.println("[INFO]    Updating total executed count for product - " + pageObjects.strProductName);
			fnUpdateExcelCellData(pageObjects.strConfluenceSheetPath, "ConfluenceResults", pageObjects.strProductName, "Executed",  String.valueOf(intTotCnt));
			System.out.println("[INFO]    Updating passed executed count for product - " + pageObjects.strProductName);
			fnUpdateExcelCellData(pageObjects.strConfluenceSheetPath, "ConfluenceResults", pageObjects.strProductName, "Passed", String.valueOf(intPassCnt));
			System.out.println("[INFO]    Updating passed % for product - " + pageObjects.strProductName);
			fnUpdateExcelCellData(pageObjects.strConfluenceSheetPath, "ConfluenceResults", pageObjects.strProductName, "Pass %", df2.format(intPassPer*100) + "%");
			System.out.println("[INFO]    Updating failed executed count for product - " + pageObjects.strProductName);
			fnUpdateExcelCellData(pageObjects.strConfluenceSheetPath, "ConfluenceResults", pageObjects.strProductName, "Failed", String.valueOf(intFailCnt));
			System.out.println("[INFO]    Updating failed % for product - " + pageObjects.strProductName);
			fnUpdateExcelCellData(pageObjects.strConfluenceSheetPath, "ConfluenceResults", pageObjects.strProductName, "Fail %", df2.format(intFailPer*100) + "%");
			
			int strRNum = fnGetExcelRowNum(pageObjects.strConfluenceSheetPath, "ConfluenceResults", "Total");
			System.out.println("[INFO]    Updating total executed count in row Total - " + pageObjects.strProductName);
			fnUpdateExcelCellData(pageObjects.strConfluenceSheetPath, "ConfluenceResults", "Total", "Executed", "SUM(B3:B" + strRNum +")");
			System.out.println("[INFO]    Updating total passed count in row Total - " + pageObjects.strProductName);
			fnUpdateExcelCellData(pageObjects.strConfluenceSheetPath, "ConfluenceResults", "Total", "Passed", "SUM(C3:C" + strRNum +")");
			System.out.println("[INFO]    Updating total failed count in row Total - " + pageObjects.strProductName);
			fnUpdateExcelCellData(pageObjects.strConfluenceSheetPath, "ConfluenceResults", "Total", "Failed", "SUM(E3:E" + strRNum +")");
			
			System.out.println("[INFO]    Reading total executed cell value from Total row");
			String strTotExecCnt = fnGetExcelData(pageObjects.strConfluenceSheetPath, "ConfluenceResults", "Total", "Executed");
			if (strTotExecCnt.trim().length() < 1) strTotExecCnt = "0";
			System.out.println("[INFO]    Reading total passed cell value from Total row");
			String strTotPassCnt = fnGetExcelData(pageObjects.strConfluenceSheetPath, "ConfluenceResults", "Total", "Passed");
			if (strTotPassCnt.trim().length() < 1) strTotPassCnt = "0";
			System.out.println("[INFO]    Reading total failed cell value from Total row");
			String strTotFailCnt = fnGetExcelData(pageObjects.strConfluenceSheetPath, "ConfluenceResults", "Total", "Failed");
			if (strTotFailCnt.trim().length() < 1) strTotFailCnt = "0";
			Double intTotPassPer = Double.parseDouble(strTotPassCnt) / Double.parseDouble(strTotExecCnt);
			Double intTotFailPer = Double.parseDouble(strTotFailCnt) / Double.parseDouble(strTotExecCnt);
			
			System.out.println("[INFO]    Updating passed % in row Total - " + pageObjects.strProductName);
			fnUpdateExcelCellData(pageObjects.strConfluenceSheetPath, "ConfluenceResults", "Total", "Pass %", String.valueOf(df2.format(intTotPassPer*100) + "%"));
			System.out.println("[INFO]    Updating failed % in row Total - " + pageObjects.strProductName);
			fnUpdateExcelCellData(pageObjects.strConfluenceSheetPath, "ConfluenceResults", "Total", "Fail %", String.valueOf(df2.format(intTotFailPer*100) + "%"));
			
			if (strProductName.trim().toLowerCase().equalsIgnoreCase(pageObjects.strProductName.trim().toLowerCase()) == false) pageObjects.strProductName = strProductName;
				
			//Steps to take the backup of the confluencedashboardresults.xls file into the results folder
			System.out.println("[INFO]    Copying the confluence updated file .....");
			FileUtils.copyFile(new File(pageObjects.strConfluenceSheetPath), new File(pageObjects.strHTMLPath + "ConfluenceDashBoardResults.xls"));

			//Launch the chrome browser
			strVerifyLoginID = false;
			pageObjects.strLoginStatus = false;
			pageObjects.intSheetRowNum = 0;
			fnLoginModule.dpmChromeBrowserLaunch("LoginSheet", "ConfluenceUpdate");
			if (fnObjectExists(objConfUserName, 15) != 0) {
				String strUserName = fnGetExcelData(pageObjects.strConfigFilepath, "LoginSheet", "ConfluenceUpdate", "UserName");
				String strPassword = fnGetExcelData(pageObjects.strConfigFilepath, "LoginSheet", "ConfluenceUpdate", "Password");
				wDriver.findElement(By.xpath(objConfUserName)).sendKeys(strUserName + Keys.TAB);
				wDriver.findElement(By.xpath(objConfPassword)).sendKeys(strPassword + Keys.TAB);
				fnObjectClick(objConfLoginBtn); }
			
			if (fnObjectExists(objConfAttachLink, 10) != 0) {
				wDriver.findElement(By.xpath(objConfAttachLink)).click();
				fnAppSync(1);
				
				if (fnObjectExists(objLockEdit, 5)!=0)
					fnObjectClick(objLockEdit);
				
				if (fnObjectExists(objConfChooseFile, 10) != 0) {
					wDriver.findElement(By.xpath(objConfChooseFile)).click();
					fnAppSync(1);
					String strTemp1 = pageObjects.strConfluenceSheetPath;
					strTemp1 = strTemp1.replace("\\\\", "\\");
					StringSelection strSelect = new StringSelection(strTemp1);
					Toolkit.getDefaultToolkit().getSystemClipboard().setContents(strSelect, null);
	
					Robot robot = new Robot();
					robot.keyPress(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_V);
					robot.keyRelease(KeyEvent.VK_V);
					robot.keyRelease(KeyEvent.VK_CONTROL);
					robot.keyPress(KeyEvent.VK_ENTER);
					robot.keyRelease(KeyEvent.VK_ENTER); 
					if (fnObjectExists(objConfAttachBtn, 2)!=0) {
						fnObjectClick(objConfAttachBtn);
						if (fnObjectExists(objUnLockEdit, 5)!=0)
							fnObjectClick(objUnLockEdit);
						if (fnObjectExists(objConfViewPageBtn, 2)!=0) 
							fnObjectClick(objConfViewPageBtn);
						fnAppSync(1);
						if (fnObjectExists(objUserMenu, 2)!=0) { 
							fnObjectClick(objUserMenu);
							if (fnObjectExists(objLogoutLnk, 4)!=0) 
								fnObjectClick(objLogoutLnk); }

						System.out.println("[INFO]   Successfully updated the test results details to CONFLUENCE......."); 
						wDriver.close();
						wDriver.quit(); } } } } }
	
	//Function to find the given string value is numeric or not
	public static boolean fnIsStringNumeric(String str) {
	  return str.matches("-?\\d+(\\.\\d+)?"); }
	
	//Function to delete the files/ folders from the given path =============================================================================
	public void fnDeleteTmpFilesFolder(String strPath, String strFileType) {
		File root = new File(strPath);
		File[] Files = root.listFiles();
		if(Files != null)
			for(File f: Files) {
				if(f.isDirectory()) {
					fnDeleteTmpFilesFolder(String.valueOf(f), strFileType);
					f.delete();	}
				else
					if (strFileType.trim().length()==0)
						f.delete();
					else {
						if (f.getName().trim().toLowerCase().contains(strFileType.trim().toLowerCase()))
							f.delete(); } }	}
	
	//function to insert the value in the given row number and column name in the excel sheet cell data ====================================
	public synchronized int fnAppendEmptyRowToExcelSheet(String fileName, String sheetName) throws Throwable {
		//to get the rows count of the excel sheet
		int strRowCnt = fnGetExcelRowCount(fileName, sheetName)+1;

		//defining the file input stream
		FileInputStream fileInputStream = new FileInputStream(new File(fileName));

		//Read the spreadsheet that needs to be updated 
		@SuppressWarnings("resource")
		HSSFWorkbook strWorkBook = new HSSFWorkbook(fileInputStream);
		
		//Access the workbook 
		HSSFSheet strworksheet = strWorkBook.getSheet(sheetName);

		//Access the worksheet, so that we can update
		HSSFRow strRow = strworksheet.createRow(strRowCnt+1);

		fileInputStream.close();
		FileOutputStream fileOutputStream = new FileOutputStream(new File(fileName));
		strWorkBook.write(fileOutputStream);
		fileOutputStream.close(); 
		System.out.println("[INFO]   Excel Sheet - " + sheetName + ", Empty row inserted at .. " + strRowCnt); 
		return strRowCnt;}

	//Function to enter the value in cell based on the given parameters ========================================================================
	@SuppressWarnings({ "deprecation", "resource" })
	public synchronized void fnInsertCellValueAtRowNum(String strFileName, String strSheetName, int rowNum, String strColName, String strCellValue, String hyperLink) throws Exception{

		//get the columns count from the excel sheet
		int colNum = fnGetExcelColNum(strFileName, strSheetName, strColName);

		//defining the file input stream
		FileInputStream fileInputStream = new FileInputStream(new File(strFileName));
		//Read the spreadsheet that needs to be updated 
		HSSFWorkbook strWorkBook = new HSSFWorkbook(fileInputStream);
		
		//Access the workbook 
		HSSFSheet strworksheet = strWorkBook.getSheet(strSheetName);

		//Access the worksheet, so that we can update
		HSSFCell strcell = strworksheet.getRow(rowNum).createCell(colNum);
		if (hyperLink.trim().toLowerCase().length()>0) {
			strcell.setCellValue(strCellValue);
			CreationHelper createHelper = strWorkBook.getCreationHelper();
			org.apache.poi.ss.usermodel.Hyperlink link = createHelper.createHyperlink(Hyperlink.LINK_FILE);
			link.setAddress(hyperLink);
			strcell.setHyperlink(link); }
		else { //update the excel cell value
			strcell.setCellValue(strCellValue); }

		fileInputStream.close();
		FileOutputStream fileOutputStream = new FileOutputStream(new File(strFileName));
		strWorkBook.write(fileOutputStream);
		fileOutputStream.close(); }

	//Step to highlight the web element ============================================================================================================
	public void highlightElement(WebDriver driver, WebElement element) {
	    for (int i = 0; i < 2; i++) {
	        JavascriptExecutor js = (JavascriptExecutor) driver;
	        js.executeScript("arguments[0].setAttribute('style', arguments[1]);",element, "color: yellow; border: 2px solid yellow;");
	        js.executeScript("arguments[0].setAttribute('style', arguments[1]);",element, ""); } }
	
	//Function to move the mouse pointer to the object ================================================================================================
	public void fnMouseMove(String strObj) throws Throwable {
		if (fnObjectExists(strObj, 1) != 0) {
		Point objCoordinates = wDriver.findElement(By.xpath(strObj)).getLocation();
		Robot robot = new Robot();
		robot.mouseMove(objCoordinates.getX()+100, objCoordinates.getY()+75);
		robot.mouseMove(objCoordinates.getX()+110, objCoordinates.getY()+85);} 	}
	
	//Function to move the mouse pointer to the object
	public void fnMouseMove1(String strObj) throws Throwable {
		if (fnObjectExists(strObj, 1) != 0) {
            Actions act = new Actions(wDriver);
            WebElement elt = wDriver.findElement(By.xpath(strObj));
            act.moveToElement(elt).build().perform();}	}
	
	//Function to select the values for the ADDITIONAL TAB - IRS/ OIS/ BASIS products =========================================
	public void fnSelectAdditionalTabDetails(String strDataSheet, String strDataPointer) throws Throwable {
		Robot robot = new Robot();
		//Fixed Pay Leg Details
		//---- PAY LAG
		String strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_PayLag");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPPayLag, 1) != 0) {
				wDriver.findElement(By.xpath(objFPPayLag)).click();
				wDriver.findElement(By.xpath(objFPPayLag)).clear();
				wDriver.findElement(By.xpath(objFPPayLag)).sendKeys(strVal);
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFPPayLag)).sendKeys(Keys.TAB); }
	
		//---- BUS/ CAL button
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_Bus_Cal");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPBUSCALBtn, 1) != 0) {
				String strVal1 = fnGetObjectAttributeValue(objFPBUSCALBtn, "outerText");
				if (strVal.trim().toLowerCase().contentEquals(strVal1.trim().toLowerCase())==false)
					fnObjectClick(objFPBUSCALBtn); }

		//---- ARR/ ADV button
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_Arr_Adv");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPARRADVBtn, 1) != 0) {
				String strVal1 = fnGetObjectAttributeValue(objFPARRADVBtn, "outerText");
				if (strVal.trim().toLowerCase().contentEquals(strVal1.trim().toLowerCase())==false)
					fnObjectClick(objFPARRADVBtn); }
		
		//---- Day Type list
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_DayType");
		int strIntTemp=0;
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPDayType, 1) != 0)
				strIntTemp = fnSelectValue(objFPDayType, strVal);
		
		//---- Adjustment Method list
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_Adj_Method");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPAdjMethod, 1) != 0)
				strIntTemp = fnSelectValue(objFPAdjMethod, strVal);

		//---- Accrual Start Date
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_Accural_StartDate");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPAccrualStartDate, 1) != 0) {
				wDriver.findElement(By.xpath(objFPAccrualStartDate)).click();
				wDriver.findElement(By.xpath(objFPAccrualStartDate)).clear();
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFPAccrualStartDate)).sendKeys(strVal);
				wDriver.findElement(By.xpath(objFPAccrualStartDate)).sendKeys(Keys.ENTER);
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFPAccrualStartDate)).sendKeys(Keys.TAB); }

		//---- Comp Method
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_CompMethod");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPCompMethod, 1) != 0)
				strIntTemp = fnSelectValue(objFPCompMethod, strVal);

		//---- Comp Frequency
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_CompFreq");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPCompFreq, 1) != 0)
				strIntTemp = fnSelectValue(objFPCompFreq, strVal);

		//---- Pay Holidays - FIXED PAY ========================================================
		int intFound = 0;
		boolean strResetOffsetHolidays = false;
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_PayHolidays");
		if (strVal.isEmpty() == false) {
			for(int intTempVal = 0; intTempVal<8; intTempVal++) {
				if (fnObjectExists(objDelPayHolidays, 1) != 0) {
					try { wDriver.findElement(By.xpath(objDelPayHolidays)).click(); }
					catch (Exception exp) {} }
				else
					intFound = 1;
				if (intFound == 1) break; }
			
			//List<WebElement> arrElement = wDriver.findElements(By.xpath(objFPPayHolidayOpenDropDown));
			//arrElement.get(0).click();
			
			//String objSelect = "//select[@id='side1_holidays.payHoliday']/option[@value='BAV']";
			//String objSelect1 = "//select[@id='side1_holidays.payHoliday']/option[@value='LON']";
			//fnObjectClick(objSelect);
			//fnObjectClick(objSelect1);

			fnObjectClick(objFPPayHolidayListOpen);
			//strVal = "BAV;LON;JMD;SEO";
			String[] strValList = strVal.split(";");
			for(int intLpCnt = 0; intLpCnt < strValList.length; intLpCnt++){
				String objTempSelect = objFPPayHolidaySelect;
				objTempSelect =  objTempSelect.replace("VALUE", strValList[intLpCnt]);
				fnObjectClick(objTempSelect);
				robot.delay(1000); }
			fnObjectClick(objFPPayHolidayLabel); }

		//Steps to select the drop down value for FP - RESET OFFSET HOLIDAYS ========================
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_ResetOffsetHolidays");
		if (strVal.isEmpty() == false) {
			List<WebElement> arrElement = wDriver.findElements(By.xpath(objDelResetOffsetHolidays));
			for (int intTempVal=0;intTempVal<=arrElement.size()+1;intTempVal++) {
				try {
				arrElement.get(intTempVal).click(); } catch(Exception exp) {}
				strResetOffsetHolidays = true; }
			
				fnObjectClick(objFPResetHolidayListOpen);
				//strVal = "BAV;LON;JMD;SEO";
				String[] strValList = strVal.split(";");
				for(int intLpCnt = 0; intLpCnt < strValList.length; intLpCnt++) {
					String objTempSelect = objFPResetHolidaySelect;
					objTempSelect =  objTempSelect.replace("VALUE", strValList[intLpCnt]);
					fnObjectClick(objTempSelect);
					robot.delay(1000); }
				fnObjectClick(objFPPayHolidayLabel);

			
			arrElement = wDriver.findElements(By.xpath(objFRResetOffsetHolidays));
			arrElement.get(0).click();
			if (fnObjectExists(objFPResetHolidayList, 1) != 0)
				strIntTemp = fnSelectValue(objFPResetHolidayList, strVal);
			fnAppSync(1);
			fnObjectClick(objCloseDropDown); 
			
		}

		//---- FP - Index Freq. Day Count
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_IndexFreqDayCount");
		if (strVal.isEmpty() == false)
			if (fnObjectExists(objFPIndexFreqDayCnt, 1) != 0)
				strIntTemp = fnSelectValue(objFPIndexFreqDayCnt, strVal);
		
		//---- FP - 1st reset rate
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_1st_ResetRate");
		if (strVal.isEmpty() == false)
			if (fnObjectExists(objFP1stResetRate, 1) != 0) {
				wDriver.findElement(By.xpath(objFP1stResetRate)).click();
				wDriver.findElement(By.xpath(objFP1stResetRate)).clear();
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFP1stResetRate)).sendKeys(strVal);
				wDriver.findElement(By.xpath(objFP1stResetRate)).sendKeys(Keys.ENTER);
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFP1stResetRate)).sendKeys(Keys.TAB); }

		//---- FP - Reset Frequency
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_ResetFreq");
		if (strVal.isEmpty() == false)
			if (fnObjectExists(objFPResetFreq, 1) != 0)
				strIntTemp = fnSelectValue(objFPResetFreq, strVal);
		
		//---- FP - reset offset
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_ResetOffset");
		if (strVal.isEmpty() == false)
			if (fnObjectExists(objFPResetOffset, 1) != 0) {
				wDriver.findElement(By.xpath(objFPResetOffset)).click();
				wDriver.findElement(By.xpath(objFPResetOffset)).clear();
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFPResetOffset)).sendKeys(strVal);
				wDriver.findElement(By.xpath(objFPResetOffset)).sendKeys(Keys.ENTER);
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFPResetOffset)).sendKeys(Keys.TAB); }
		
		//---- ARR/ ADV button
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_Arr_Adv2");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPBasicArrAdvBtn, 1) != 0) {
				String strVal1 = fnGetObjectAttributeValue(objFPBasicArrAdvBtn, "outerText");
				if (strVal.trim().toLowerCase().contentEquals(strVal1.trim().toLowerCase())==false)
					fnObjectClick(objFPBasicArrAdvBtn); }
		//======================================================================================
		//---- Stub Mode - 1
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_StubMode1");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPStubMode1, 1) != 0)
				strIntTemp = fnSelectValue(objFPStubMode1, strVal);
		
		//---- Rate
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_Rate");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPRate, 1) != 0) {
				wDriver.findElement(By.xpath(objFPRate)).click();
				wDriver.findElement(By.xpath(objFPRate)).clear();
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFPRate)).sendKeys(strVal);
				wDriver.findElement(By.xpath(objFPRate)).sendKeys(Keys.ENTER);
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFPRate)).sendKeys(Keys.TAB); }
		
		//---- Day Count
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_DayCount");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPDayCount, 1) != 0)
				strIntTemp = fnSelectValue(objFPDayCount, strVal);

		//---- Stub Mode - 2
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_StubMode2");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPStubMode2, 1) != 0)
				strIntTemp = fnSelectValue(objFPStubMode2, strVal);

		//---- 1st Pay Date - 1
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_1stPayDate1");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPFirstPayDate1, 1) != 0) {
				wDriver.findElement(By.xpath(objFPFirstPayDate1)).click();
				wDriver.findElement(By.xpath(objFPFirstPayDate1)).clear();
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFPFirstPayDate1)).sendKeys(strVal);
				wDriver.findElement(By.xpath(objFPFirstPayDate1)).sendKeys(Keys.ENTER);
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFPFirstPayDate1)).sendKeys(Keys.TAB); }
		
		//---- 1st Roll Date - 1
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_1stRollDate1");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPFirstRollDate1, 1) != 0) {
				wDriver.findElement(By.xpath(objFPFirstRollDate1)).click();
				wDriver.findElement(By.xpath(objFPFirstRollDate1)).clear();
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFPFirstRollDate1)).sendKeys(strVal);
				wDriver.findElement(By.xpath(objFPFirstRollDate1)).sendKeys(Keys.ENTER);
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFPFirstRollDate1)).sendKeys(Keys.TAB); }
		
		//---- 1st Pay Date - 2
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_1stPayDate2");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPFirstPayDate2, 1) != 0) {
				wDriver.findElement(By.xpath(objFPFirstPayDate2)).click();
				wDriver.findElement(By.xpath(objFPFirstPayDate2)).clear();
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFPFirstPayDate2)).sendKeys(strVal);
				wDriver.findElement(By.xpath(objFPFirstPayDate2)).sendKeys(Keys.ENTER);
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFPFirstPayDate2)).sendKeys(Keys.TAB); }
		
		//---- 1st Roll Date - 2
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_1stRollDate2");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPFirstRollDate2, 1) != 0) {
				wDriver.findElement(By.xpath(objFPFirstRollDate2)).click();
				wDriver.findElement(By.xpath(objFPFirstRollDate2)).clear();
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFPFirstRollDate2)).sendKeys(strVal);
				wDriver.findElement(By.xpath(objFPFirstRollDate2)).sendKeys(Keys.ENTER);
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFPFirstRollDate2)).sendKeys(Keys.TAB); }

		//=================================== FLOAT RECEIVE =============================================================

		//Float Receive Details
		//---- PAY LAG
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FR_PayLag");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRPayLag, 1) != 0) {
				wDriver.findElement(By.xpath(objFRPayLag)).click();
				wDriver.findElement(By.xpath(objFRPayLag)).clear();
				wDriver.findElement(By.xpath(objFRPayLag)).sendKeys(strVal);
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFRPayLag)).sendKeys(Keys.TAB); }
	
		//---- BUS/ CAL button
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FR_Bus_Cal");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRBUSCALBtn, 1) != 0) {
				String strVal1 = fnGetObjectAttributeValue(objFRBUSCALBtn, "outerText");
				if (strVal.trim().toLowerCase().contentEquals(strVal1.trim().toLowerCase())==false)
					fnObjectClick(objFRBUSCALBtn); }

		//---- ARR/ ADV button
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FR_Arr_Adv1");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRARRADVBtn, 1) != 0) {
				String strVal1 = fnGetObjectAttributeValue(objFRARRADVBtn, "outerText");
				if (strVal.trim().toLowerCase().contentEquals(strVal1.trim().toLowerCase())==false)
					fnObjectClick(objFRARRADVBtn); }

		//---- Day Type list
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FR_DayType");
		strIntTemp=0;
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRDayType, 1) != 0)
				strIntTemp = fnSelectValue(objFRDayType, strVal);
		
		//---- Adjustment Method list
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FR_Adj_Method");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRAdjMethod, 1) != 0)
				strIntTemp = fnSelectValue(objFRAdjMethod, strVal);
				
		//---- Accrual Start Date
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FR_Accural_StartDate");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRAccrualStartDate, 1) != 0) {
				wDriver.findElement(By.xpath(objFRAccrualStartDate)).click();
				wDriver.findElement(By.xpath(objFRAccrualStartDate)).clear();
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFRAccrualStartDate)).sendKeys(strVal);
				wDriver.findElement(By.xpath(objFRAccrualStartDate)).sendKeys(Keys.ENTER);
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFRAccrualStartDate)).sendKeys(Keys.TAB); }

		//---- Comp Method
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FR_CompMethod");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRCompMethod, 1) != 0)
				strIntTemp = fnSelectValue(objFRCompMethod, strVal);
		
		//---- SPREAD - FLOAT RECEIVE
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FR_Spread");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRSpread, 1) != 0) {
				wDriver.findElement(By.xpath(objFRSpread)).click();
				wDriver.findElement(By.xpath(objFRSpread)).clear();
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFRSpread)).sendKeys(strVal);
				wDriver.findElement(By.xpath(objFRSpread)).sendKeys(Keys.ENTER);
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFRSpread)).sendKeys(Keys.TAB); }

		//---- Pay Holidays/ RESET OFFSET HOLIDAYS - FLOAT RECEIVE =============================
		
		intFound = 0;
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FR_PayHolidays");
		if (strVal.isEmpty() == false) {	
			
			fnObjectClick(objFRPayHolidayListOpen);
			//strVal = "BAV;LON;JMD;SEO";
			String[] strValList = strVal.split(";");
			for(int intLpCnt = 0; intLpCnt < strValList.length; intLpCnt++){
				String objTempSelect = objFRPayHolidaySelect;
				objTempSelect =  objTempSelect.replace("VALUE", strValList[intLpCnt]);
				fnObjectClick(objTempSelect);
				robot.delay(1000); }
			fnObjectClick(objFPPayHolidayLabel);	
		}

		
		if (strVal.isEmpty() == false) {
			for(int intTempVal = 0; intTempVal<5; intTempVal++) {
				if (fnObjectExists(objDelPayHolidays, 1) != 0) {
					fnObjectClick(objDelPayHolidays); }
				else
					intFound = 1;
				if (intFound == 1) break; } 
			List<WebElement> arrElement = wDriver.findElements(By.xpath(objFPPayHolidayDropDown));
			arrElement.get(1).click();
			if (fnObjectExists(objFRPayHolidayList, 1) != 0)
				strIntTemp = fnSelectValue(objFRPayHolidayList, strVal);
			fnAppSync(2);
			fnObjectClick(objFPLabel); }
	

		// --- FLOAT RECEIVE - RESET OFFSET HOLIDAY =========================================
		intFound = 0;
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FR_ResetOffsetHolidays");
		if (strVal.isEmpty() == false) {
			List<WebElement> arrElement = wDriver.findElements(By.xpath(objDelResetOffsetHolidays));
			if (strResetOffsetHolidays == false) {
				for (int intTempVal=0;intTempVal<=arrElement.size();intTempVal++)
					try {
					arrElement.get(0).click(); } catch(Exception exp) {} }
			
			fnObjectClick(objFRResetHolidayListOpen);
			//strVal = "BAV;LON;JMD;SEO";
			String[] strValList = strVal.split(";");
			for(int intLpCnt = 0; intLpCnt < strValList.length; intLpCnt++) {
				String objTempSelect = objFRResetHolidaySelect;
				objTempSelect =  objTempSelect.replace("VALUE", strValList[intLpCnt]);
				fnObjectClick(objTempSelect);
				robot.delay(1000); }
			fnObjectClick(objFPPayHolidayLabel);

			for(int intTempVal = 0; intTempVal<5; intTempVal++) {
				if (fnObjectExists(objDelResetOffsetHolidays, 1) != 0) {
					fnObjectClick(objDelResetOffsetHolidays); }
				else
					intFound = 1;
				if (intFound == 1) break; }
			arrElement = wDriver.findElements(By.xpath(objFRResetOffsetHolidays));
			int maxCnt = arrElement.size();
			arrElement.get(maxCnt-1).click();
			//fnObjectClick(objFRResetOffsetHolidays);
			if (fnObjectExists(objFRResetHolidayList, 1) != 0)
				strIntTemp = fnSelectValue(objFRResetHolidayList, strVal);
			fnAppSync(2);
			fnObjectClick(objFPLabel); 
			
			}
		
		//======================================================================================
		//---- Index Freq Day Count
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FR_IndexFreqDayCount");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRIndexFreqDayCount, 1) != 0)
				strIntTemp = fnSelectValue(objFRIndexFreqDayCount, strVal);

		//---- 1st RESET RATE
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FR_1st_ResetRate");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRFirstResetRate, 1) != 0) {
				wDriver.findElement(By.xpath(objFRFirstResetRate)).click();
				wDriver.findElement(By.xpath(objFRFirstResetRate)).clear();
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFRFirstResetRate)).sendKeys(strVal);
				wDriver.findElement(By.xpath(objFRFirstResetRate)).sendKeys(Keys.ENTER);
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFRFirstResetRate)).sendKeys(Keys.TAB); }

		//---- Reset Freq.
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FR_ResetFreq");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRResetFreq, 1) != 0)
				strIntTemp = fnSelectValue(objFRResetFreq, strVal);
		
		//---- 1st RESET OFFSET
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FR_ResetOffset");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRResetOffset, 1) != 0) {
				wDriver.findElement(By.xpath(objFRResetOffset)).click();
				wDriver.findElement(By.xpath(objFRResetOffset)).clear();
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFRResetOffset)).sendKeys(strVal);
				wDriver.findElement(By.xpath(objFRResetOffset)).sendKeys(Keys.ENTER);
				robot.delay(1000);
				wDriver.findElement(By.xpath(objFRResetOffset)).sendKeys(Keys.TAB); }
		
		//---- ARR/ ADV button
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FR_Arr_Adv2");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRADVARRBtn, 1) != 0) {
				String strVal1 = fnGetObjectAttributeValue(objFRADVARRBtn, "outerText");
				if (strVal.trim().toLowerCase().contentEquals(strVal1.trim().toLowerCase())==false)
					fnObjectClick(objFRADVARRBtn); } }

	//Function to validate the values on the ADDITIONAL TAB - IRS/ OIS/ BASIS products =========================================
	public void fnValidateAdditionalTabDetails(String strDataSheet, String strDataPointer) throws Throwable {

		//Step to find out the which Index Freq is greater
		String strIndxFrq1 = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "IndexFreq1").trim().toLowerCase();
		String strIndxFrq2 = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "IndexFreq2").trim().toLowerCase();
		int intValidateSide = 1;
		String strValidateSide = "";
		//System.out.println("Indx 1 - " + strIndxFrq1 + ", Indx 2 - " + strIndxFrq2);
		if (pageObjects.strSaveAs == false) {
			if (strIndxFrq1.contentEquals("1m")) {
				if (strIndxFrq2.contentEquals("1d"))
					intValidateSide = 2; }
			else if (strIndxFrq1.contentEquals("3m")) {
				if (strIndxFrq2.contentEquals("1d") || strIndxFrq2.contentEquals("1m"))
					intValidateSide = 2; }
			else if (strIndxFrq1.contentEquals("6m")) {
				if (strIndxFrq2.contentEquals("1d") || strIndxFrq2.contentEquals("1m") || strIndxFrq2.contentEquals("3m"))
					intValidateSide = 2; }
			else if (strIndxFrq1.contentEquals("1y")) {
				if (strIndxFrq2.contentEquals("1d") || strIndxFrq2.contentEquals("1m") || strIndxFrq2.contentEquals("3m") || strIndxFrq2.contentEquals("6m"))
					intValidateSide = 2; } }
		
		//Fixed Pay Leg Details
		//---- PAY LAG
		strValidateSide = "AD_FP_PayLag";
		if (intValidateSide==2) strValidateSide = "AD_FR_PayLag";
		String strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPPayLag, 1) != 0) {
				//step to retrieve the COUNTER PARTY from excel and book detail screen and compare
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - Pay Lag", strValidateSide, objFPPayLag, "value");	}

		//---- BUS/ CAL button
		strValidateSide = "AD_FP_Bus_Cal";
		if (intValidateSide==2) strValidateSide = "AD_FR_Bus_Cal";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPBUSCALBtn, 1) != 0)
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - BUS/CALL", strValidateSide, objFPBUSCALBtn, "outerText");

		//---- BUS/ CAL button
		strValidateSide = "AD_FP_Arr_Adv";
		if (intValidateSide==2) strValidateSide = "AD_FR_Arr_Adv1";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPBUSCALBtn, 1) != 0)
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - ARR/ADV", strValidateSide, objFPARRADVBtn, "outerText");

		//---- Day Type list
		int strIntTemp=0;
		strValidateSide = "AD_FP_DayType";
		if (intValidateSide==2) strValidateSide = "AD_FR_DayType";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPDayType, 1) != 0)
				fnListObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - Day Type", strValidateSide, objFPDayType, "value");
				
		//---- Adjustment Method list
		strValidateSide = "AD_FP_Adj_Method";
		if (intValidateSide==2) strValidateSide = "AD_FR_Adj_Method";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPAdjMethod, 1) != 0)
				fnListObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - Day Type", strValidateSide, objFPAdjMethod, "value");
		
		//---- Accrual Start Date
		strValidateSide = "AD_FP_Accural_StartDate";
		if (intValidateSide==2) strValidateSide = "AD_FR_Accural_StartDate";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPAccrualStartDate, 1) != 0)
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - Accural Start Date", strValidateSide, objFPAccrualStartDate, "value");

		//---- Comp Method
		strValidateSide = "AD_FP_CompMethod";
		if (intValidateSide==2) strValidateSide = "AD_FR_CompMethod";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPCompMethod, 1) != 0)
				strIntTemp = fnSelectValue(objFPCompMethod, strVal);

		//---- Comp Frequency
		//int intFound = 0;
		//boolean strResetOffsetHolidays = false;
		strValidateSide = "AD_FP_CompFreq";
		if (intValidateSide==2) strValidateSide = "AD_FR_CompFreq";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPCompFreq, 1) != 0)
				fnListObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - Comp. Freq.", strValidateSide, objFPCompFreq, "value");
	
		//---- Pay Holidays - FIXED PAY ========================================================
		strValidateSide = "AD_FP_PayHolidays";
		if (intValidateSide==2) strValidateSide = "AD_FR_PayHolidays";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		String[] strValList = strVal.split(";");
		
		for(int intTemp=0;intTemp<strValList.length;intTemp++) {
			if (fnObjectExists(objPayHolidays1, 1)!=0)
			if (strVal.isEmpty() == false) {
				//strVal = strVal.trim().substring(0,  3);
				strVal = strValList[intTemp];
				String strVal2="";
				strVal2 = fnGetObjectAttributeValue(objPayHolidays1, "outerText");
				if (strVal2.trim().toLowerCase().contentEquals("null")==false){				
					if(strVal2.trim().toLowerCase().contains(strVal.trim().toLowerCase()) || strVal.trim().toLowerCase().contains(strVal2.trim().toLowerCase())) {
						System.out.println("[INFO]    == == == List String Comparison PASSED - FP - Pay Holidays - Actual Value : " + strVal2 + " - Expected Value : " + strVal); }
						else {
							System.out.println("[ERROR]    ****** List String Comparison FAILED - FP - Pay Holidays - Actual Value : " + strVal2 + " - Expected Value : " + strVal);
							pageObjects.strCompareVal = false;
							if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
								pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + " Field name : FP - Pay Holidays - Actual Value : " + strVal2 + " - Expected Value : " + strVal + "; \n "; }	} }	}

		//Steps to select the drop down value for RESET OFFSET HOLIDAYS ========================
		strValidateSide = "AD_FP_ResetOffsetHolidays";
		if (intValidateSide==2) strValidateSide = "AD_FR_ResetOffsetHolidays";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		strValList = strVal.split(";");

		//String strObj = objSelectAssetHolidays1;
		//if (strDataSheet.trim().toLowerCase().contains("defaultvalues") || (strSaveAs == true))
		
		for(int intTemp=0;intTemp<strValList.length;intTemp++){
			if (fnObjectExists(objOffsetHolidays1, 1)!=0)
			if (strVal.isEmpty() == false) {
				//strVal = strVal.trim().substring(0,  3);
				strVal = strValList[intTemp];
				String strVal2="";
				strVal2 = fnGetObjectAttributeValue(objOffsetHolidays1, "outerText");
				if (strVal2.trim().toLowerCase().contentEquals("null")==false){		
					if(strVal2.trim().toLowerCase().contains(strVal.trim().toLowerCase()) || strVal.trim().toLowerCase().contains(strVal2.trim().toLowerCase())) {
						System.out.println("[INFO]    == == == List String Comparison PASSED - FP - Reset offset Holidays - Actual Value : " + strVal2 + " - Expected Value : " + strVal); }
						else {
							System.out.println("[ERROR]    ****** List String Comparison FAILED - FP - Reset offset Holidays - Actual Value : " + strVal2 + " - Expected Value : " + strVal);
							pageObjects.strCompareVal = false;
							if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
								pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + " Field name : FP - Reset offset Holidays - Actual Value : " + strVal2 + " - Expected Value : " + strVal + "; \n "; } } } }

		//======================================================================================
		//---- FP - Index Freq. Day Count
		strValidateSide = "AD_FP_IndexFreqDayCount";
		if (intValidateSide==2) strValidateSide = "AD_FR_IndexFreqDayCount";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPIndexFreqDayCnt, 1) != 0)
				fnListObjCompareValues(objFPIndexFreqDayCnt, strDataSheet, strDataPointer, "Float Pay - Index Frequency Day Count", strValidateSide, objFPIndexFreqDayCnt, "value");

		//---- FP - 1st reset rate
		strValidateSide = "AD_FP_1st_ResetRate";
		if (intValidateSide==2) strValidateSide = "AD_FR_1st_ResetRate";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFP1stResetRate, 1) != 0)
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - 1st reset rate", strValidateSide, objFP1stResetRate, "value");
		
		//---- FP - Reset Frequency
		strValidateSide = "AD_FP_ResetFreq";
		if (intValidateSide==2) strValidateSide = "AD_FP_ResetFreq";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPResetFreq, 1) != 0)
				fnListObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - Reset Freq", strValidateSide, objFPResetFreq, "value");

		//---- FP - Reset Offset
		strValidateSide = "AD_FP_ResetOffset";
		if (intValidateSide==2) strValidateSide = "AD_FR_ResetOffset";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPResetOffset, 1) != 0)
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - Reset Offset", strValidateSide, objFPResetOffset, "value");
		
		//---- ARR/ ADV button
		strValidateSide = "AD_FP_Arr_Adv2";
		if (intValidateSide==2) strValidateSide = "AD_FR_Arr_Adv2";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPBasicArrAdvBtn, 1) != 0)
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - ARR/ ADV", strValidateSide, objFPBasicArrAdvBtn, "outerText");
		
		//======================================================================================
		//---- Stub Mode - 1
		strValidateSide = "AD_FP_StubMode1";
		if (intValidateSide==2) strValidateSide = "AD_FP_StubMode2";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPStubMode1, 1) != 0)
				fnListObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - Stub Mode 1", strValidateSide, objFPStubMode1, "value");
		
		//---- Rate
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_Rate");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPRate, 1) != 0)
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - Rate", "AD_FP_Rate", objFPRate, "outerText");
		
		//---- Day Count
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_DayCount");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPDayCount, 1) != 0)
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - Day Count", "AD_FP_DayCount", objFPDayCount, "value");

		//---- Stub Mode - 2
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_StubMode2");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPStubMode2, 1) != 0)
				fnListObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - Stub Mode 2", "AD_FP_StubMode2", objFPStubMode2, "value");

		//---- 1st Pay Date - 1
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_1stPayDate1");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPFirstPayDate1, 1) != 0) 
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - 1st Pay Date - 1", "AD_FP_1stPayDate1", objFPFirstPayDate1, "value");

		//---- 1st Roll Date - 1
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_1stRollDate1");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPFirstRollDate1, 1) != 0) 
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - 1st Roll Date - 1", "AD_FP_1stRollDate1", objFPFirstRollDate1, "value");
		
		//---- 1st Pay Date - 2
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_1stPayDate2");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPFirstPayDate2, 1) != 0) 
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - 1st Pay Date - 2", "AD_FP_1stPayDate2", objFPFirstPayDate2, "value");
		
		//---- 1st Roll Date - 2
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "AD_FP_1stRollDate2");
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFPFirstRollDate2, 1) != 0) 
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Pay - 1st Roll Date -2 ", "AD_FP_1stRollDate2", objFPFirstRollDate2, "value");

		//=================================== FLOAT RECEIVE =============================================================
		//Float Receive Details
		//---- PAY LAG
		strValidateSide = "AD_FR_PayLag";
		if (intValidateSide==2) strValidateSide = "AD_FP_PayLag";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRPayLag, 1) != 0) 
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Receive - PayLag", strValidateSide, objFRPayLag, "value");
	
		//---- BUS/ CAL button
		strValidateSide = "AD_FR_Bus_Cal";
		if (intValidateSide==2) strValidateSide = "AD_FP_Bus_Cal";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRBUSCALBtn, 1) != 0)
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Receive - BUS/ CAL", strValidateSide, objFRBUSCALBtn, "outerText");

		//---- ARR/ ADV button
		strValidateSide = "AD_FR_Arr_Adv1";
		if (intValidateSide==2) strValidateSide = "AD_FP_Arr_Adv1";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRARRADVBtn, 1) != 0) 
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Receive - ARR/ ADV", strValidateSide, objFRARRADVBtn, "outerText");

		//---- Day Type list
		strValidateSide = "AD_FR_DayType";
		if (intValidateSide==2) strValidateSide = "AD_FP_DayType";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		strIntTemp=0;
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRDayType, 1) != 0)
				fnListObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Receive - Day Type", strValidateSide, objFRDayType, "value");
		
		//---- Adjustment Method list
		strValidateSide = "AD_FR_Adj_Method";
		if (intValidateSide==2) strValidateSide = "AD_FP_Adj_Method";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRAdjMethod, 1) != 0)
				fnListObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Receive - Adjustment Method", strValidateSide, objFRAdjMethod, "value");
				
		//---- Accrual Start Date
		strValidateSide = "AD_FR_Accural_StartDate";
		if (intValidateSide==2) strValidateSide = "AD_FP_Accural_StartDate";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRAccrualStartDate, 1) != 0) 
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Receive - Accural Start Date", strValidateSide, objFRAccrualStartDate, "value");

		//---- Comp Method
		strValidateSide = "AD_FR_CompMethod";
		if (intValidateSide==2) strValidateSide = "AD_FP_CompMethod";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRCompMethod, 1) != 0)
				fnListObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Receive - Comp. Method", strValidateSide, objFRCompMethod, "value");
		
		//---- SPREAD - FLOAT RECEIVE
		strValidateSide = "AD_FR_Spread";
		if (intValidateSide==2) strValidateSide = "AD_FP_Spread";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRSpread, 1) != 0) 
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Receive - Spread", strValidateSide, objFRSpread, "value");

		//---- Pay Holidays/ RESET OFFSET HOLIDAYS - FLOAT RECEIVE =============================
		strValidateSide = "AD_FR_PayHolidays";
		if (intValidateSide==2) strValidateSide = "AD_FP_PayHolidays";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		strValList = strVal.split(";");
		for(int intTemp=0;intTemp<strValList.length;intTemp++){
			if (fnObjectExists(objPayHolidays2, 1) != 0) 
			if (strVal.isEmpty() == false) {
				//strVal = strVal.trim().substring(0,  3);
				strVal = strValList[intTemp];
				String strVal2="";
				strVal2 = fnGetObjectAttributeValue(objPayHolidays2, "outerText");
				if (strVal2.trim().toLowerCase().contentEquals("null")==false) {
					if(strVal2.trim().toLowerCase().contains(strVal.trim().toLowerCase()) || strVal.trim().toLowerCase().contains(strVal2.trim().toLowerCase())) {
						System.out.println("[INFO]    == == == List String Comparison PASSED - FR - Pay Holidays - Actual Value : " + strVal2 + " - Expected Value : " + strVal); }
						else {
							System.out.println("[ERROR]    ****** List String Comparison FAILED - FR - Pay Holidays - Actual Value : " + strVal2 + " - Expected Value : " + strVal);
							pageObjects.strCompareVal = false;
							if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
								pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + " Field name : FR - Pay Holidays - Actual Value : " + strVal2 + " - Expected Value : " + strVal + "; \n "; }	} }	}

		// --- FLOAT RECEIVE - RESET OFFSET HOLIDAY =========================================
		strValidateSide = "AD_FR_ResetOffsetHolidays";
		if (intValidateSide==2) strValidateSide = "AD_FP_ResetOffsetHolidays";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		strValList = strVal.split(";");
		for(int intTemp=0;intTemp<strValList.length;intTemp++){
			if (strVal.isEmpty() == false) {
				//strVal = strVal.trim().substring(0,  3);
				String strVal3 = strValList[intTemp];
				String strHiddenVal="", strVal2="";
				strVal2 = fnGetObjectAttributeValue(objOffsetHolidays2, "outerText");
				if (strVal2.trim().toLowerCase().contentEquals("null")==false) {
					if(strVal2.trim().toLowerCase().contains(strVal3.trim().toLowerCase()) || strVal3.trim().toLowerCase().contains(strVal2.trim().toLowerCase())) {
						System.out.println("[INFO]    == == == List String Comparison PASSED - FP - Reset offset Holidays - Actual Value : " + strVal2 + " - Expected Value : " + strVal); }
						else {
							System.out.println("[ERROR]    ****** List String Comparison FAILED - FP - Reset offset Holidays - Actual Value : " + strVal2 + " - Expected Value : " + strVal);
							pageObjects.strCompareVal = false;
							if (pageObjects.strGVCompareValues.isEmpty()) pageObjects.strGVCompareValues = "";
								pageObjects.strGVCompareValues =  pageObjects.strGVCompareValues + " Field name : FP - Reset offset Holidays - Actual Value : " + strVal2 + " - Expected Value : " + strVal + "; \n "; } } }	}
		//======================================================================================
		//---- Index Freq Day Count
		strValidateSide = "AD_FR_IndexFreqDayCount";
		if (intValidateSide==2) strValidateSide = "AD_FP_IndexFreqDayCount";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRIndexFreqDayCount, 1) != 0)
				fnListObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Receive - Index Freq Day Count", strValidateSide, objFRIndexFreqDayCount, "value");

		//---- 1st RESET RATE
		strValidateSide = "AD_FR_1st_ResetRate";
		if (intValidateSide==2) strValidateSide = "AD_FP_1st_ResetRate";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRFirstResetRate, 1) != 0)
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Receive - 1st Reset Rate", strValidateSide, objFRFirstResetRate, "value");

		//---- Reset Freq.
		strValidateSide = "AD_FR_ResetFreq";
		if (intValidateSide==2) strValidateSide = "AD_FP_ResetFreq";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRResetFreq, 1) != 0)
				fnListObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Receive - Reset Freq.", strValidateSide, objFRResetFreq, "value");
		
		//---- 1st RESET OFFSET
		strValidateSide = "AD_FR_ResetOffset";
		if (intValidateSide==2) strValidateSide = "AD_FP_ResetOffset";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRResetOffset, 1) != 0) 
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Receive - Reset Offset", strValidateSide, objFRResetOffset, "value");
		
		//---- ARR/ ADV button
		strValidateSide = "AD_FR_Arr_Adv2";
		if (intValidateSide==2) strValidateSide = "AD_FP_Arr_Adv2";
		strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strValidateSide);
		if (strVal.isEmpty() == false) 
			if (fnObjectExists(objFRADVARRBtn, 1) != 0) 
				fnStringObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "Float Receive - ARR/ ADV - 2", strValidateSide, objFRADVARRBtn, "outerText");	}

//Function to perform the SaveAs functionality for the new deal =========================================================================================================
	public void fnSaveAsDealDetails(String strDataSheet, String strDataPointer) throws Throwable {
		String strTradeFolderName ;
		String strTradeName;
		//String strCurrencyVal = "";
		String strTxt = "";
		String strProdName = "";
		Robot robot = new Robot();

		//strCurrencyVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "Currency");
		if (strDataSheet.trim().toLowerCase().contains("basis")) strProdName = "b";
		else if (strDataSheet.trim().toLowerCase().contains("irs")) strProdName = "i";
		else if (strDataSheet.trim().toLowerCase().contains("ois")) strProdName = "o";
		else if (strDataSheet.trim().toLowerCase().contains("eqo")) strProdName = "e";
		else if (strDataSheet.trim().toLowerCase().contains("loan")) strProdName = "l";
		else if (strDataSheet.trim().toLowerCase().contains("deposit")) strProdName = "d";

		//Trade folder name and Trade name generated
		int hours = LocalDateTime.now().getHour();
		int minutes = LocalDateTime.now().getMinute();
		int seconds = LocalDateTime.now().getSecond();
		
		strTradeFolderName = "AUT_" + pageObjects.strGVDate + "_" + pageObjects.strEnvironment;
		strTradeName = strProdName + String.valueOf(hours) + String.valueOf(minutes) + String.valueOf(seconds);
		
		fnUpdateExcelCellData(strDataSheetPath, strDataSheet, strDataPointer, "DPM_Deal_Name", strTradeName);
		fnUpdateExcelCellData(strDataSheetPath, strDataSheet, strDataPointer, "DPM_Deal_FolderName", strTradeFolderName);
		
		//Click on SaveAs button
		if (fnObjectExists(objDealCancelBtn, 2) != 0) {
			if (fnObjectExists(objBtnSaveAs, 2) != 0) {
				pageObjects.strGVStepName = "Click on the SaveAs button";
				fnScreenCapture(wDriver, strScreenPath.concat("dpmDealSaveAsBtn") + "-"+ strDataPointer);
				fnObjectClick(objBtnSaveAs); }
			else if (fnObjectExists(objBtnCopy, 2) != 0) {
				pageObjects.strGVStepName = "Click on the Copy button";
				fnScreenCapture(wDriver, strScreenPath.concat("dpmDealCopyBtn") + "-"+ strDataPointer);
				fnObjectClick(objBtnCopy);
				robot.delay(1000);
				if (fnObjectExists(objConfirmKillAction1, 2) != 0) {
					pageObjects.strGVStepName = "Click on the CONFIRM button";
					fnScreenCapture(wDriver, strScreenPath.concat("dpmDealSaveAsBtn") + "-"+ strDataPointer);
					fnObjectClick(objConfirmKillAction1);
					
					if (fnObjectExists(objBtnSaveAs, 5) != 0) {
						pageObjects.strGVStepName = "Click on the SaveAs button";
						fnScreenCapture(wDriver, strScreenPath.concat("dpmDealSaveAsBtn") + "-"+ strDataPointer);
						fnObjectClick(objBtnSaveAs); } } }

			//Enter Trade Folder name
			List<WebElement> eleObjArray = null, eleObjArray1 = null, eleObjArray2 = null;
			eleObjArray = wDriver.findElements(By.xpath(objSaveAsTradeFolder));
			eleObjArray1 = wDriver.findElements(By.xpath(objSaveAsTradeName));
			eleObjArray2 = wDriver.findElements(By.xpath(objSaveAsConfirmAction));
			int intArrSize = eleObjArray.size()-1;
			
			if (eleObjArray.size() > 0) {
				eleObjArray.get(intArrSize).click();
				eleObjArray.get(intArrSize).clear();
				eleObjArray.get(intArrSize).sendKeys("");
				eleObjArray.get(intArrSize).sendKeys(strTradeFolderName);
				robot.delay(1000);
				eleObjArray.get(intArrSize).sendKeys(Keys.ENTER);
				robot.delay(1000);

				//Enter Trade name
				if (eleObjArray1.size()>0) {
					eleObjArray1.get(intArrSize).click();
					eleObjArray1.get(intArrSize).clear();
					eleObjArray1.get(intArrSize).sendKeys("");
					eleObjArray1.get(intArrSize).sendKeys(strTradeName);
					robot.delay(1000);
					eleObjArray1.get(intArrSize).sendKeys(Keys.ENTER);
					robot.delay(1000); } }

			//Click on confirmation button
			fnAppSync(1);
			if (eleObjArray2.size()>0) {
				pageObjects.strGVStepName = "Click on the SaveAs Confirm button";
				fnScreenCapture(wDriver, strScreenPath.concat("dpmDealConfirmBtn") + "-"+ strDataPointer);
				eleObjArray2.get(intArrSize).click();
				//fnObjectClick(objSaveAsConfirmAction);
				robot.delay(1000); 
	
				int strFound = -1;
				boolean strSTxt = false;
				for(int intTemp = 0; intTemp<20; intTemp++) {
					strFound = -1;
					if ((fnObjectExists(objIDSuccessText, 1) != 0) || (fnObjectExists(objIDInBookingText, 1) != 0))
					{ 	String strTxt1="";
						if (fnObjectExists(objIDSuccessText, 1) != 0){
							Point objCoordinates = wDriver.findElement(By.xpath(objIDSuccessText)).getLocation();
							robot = new Robot();
							robot.mouseMove(objCoordinates.getX()+66, objCoordinates.getY()+86);
							strTxt = wDriver.findElement(By.xpath(objIDContent)).getText();
							strSTxt = (wDriver.findElement(By.xpath(objIDSuccessText)).getText().toLowerCase().contains("success"));
							robot.mouseMove(objCoordinates.getX()+65, objCoordinates.getY()+84);
							strTxt1 = wDriver.findElement(By.xpath(objIDContent)).getText(); }
						else if (fnObjectExists(objIDInBookingText, 1) != 0){
							Point objCoordinates = wDriver.findElement(By.xpath(objIDInBookingText)).getLocation();
							robot = new Robot();
							robot.mouseMove(objCoordinates.getX()+66, objCoordinates.getY()+86);
							strTxt = wDriver.findElement(By.xpath(objIDInBookingContent)).getText();
							strSTxt = (wDriver.findElement(By.xpath(objIDInBookingText)).getText().toLowerCase().contains("in booking"));
							robot.mouseMove(objCoordinates.getX()+65, objCoordinates.getY()+84);
							strTxt1 = wDriver.findElement(By.xpath(objIDInBookingContent)).getText(); }
							if (strTxt1.trim().toLowerCase().contains("object") == true) {
								strTxt = "[Object Object]";
								strFound = 0; }
							else 
								strFound = 1; }
					else if (fnObjectExists(objIDErrText, 1) != 0)
					{	Point objCoordinates = wDriver.findElement(By.xpath(objIDErrText)).getLocation();
						robot = new Robot();
						robot.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+87);
						strTxt = wDriver.findElement(By.xpath(objServerErrorContent)).getText();
						strSTxt = (wDriver.findElement(By.xpath(objServerErrorContent)).getText().toLowerCase().contains("success"));
						robot.mouseMove(objCoordinates.getX()+64, objCoordinates.getY()+83);
						strFound = 0; pageObjects.strScnStatus = false; }
					else if (fnObjectExists(objFormErr, 1) != 0)
					{	strFound = 0;
						strTxt = fnGetObjectAttributeValue(objFormErr, "outerText");
						pageObjects.strScnStatus = false; }
					robot.delay(1000);
					if (strFound != -1) break;}
		
				if (strFound == 1) {
					System.out.println("[INFO]   Deal details Popup Message details - " + strTxt);
					if (strSTxt == true) {
						if (pageObjects.strDealAmend == false) {
							pageObjects.strGVStepName = "Deal Details successfully saved with Deal Name - " + strTxt;
							System.out.println("[INFO]   Deal Details successfully saved with Deal Name - " + strTxt); } 
		
						pageObjects.strGVDealID = strTxt;
						fnScreenCapture(wDriver, strScreenPath.concat("dpmDealSaveAsSuccess") + "-"+ strDataPointer);
						fnObjectClick(objPXVLabel);
						if (fnObjectExists(objIDSuccessText, 2)!=1) {
							fnObjectClick(objIDSuccessText);
							fnObjectClick(objIDSuccessText); }
						fnAppSync(1);
						//step to update the Trade folder name and Trade name
						String strTime = fnGetCurrentDate("yyyy/mm/dd") + " " + fnGetCurrentTime();
						fnUpdateExcelCellData(strDataSheetPath, strDataSheet, strDataPointer, "Deal_DateTime", strTime); }
					else
						strFound = 0; }

				if (strFound == 0) {
					pageObjects.strScnStatus = false;
					pageObjects.strGVErrMsg = "Failed to save the " + pageObjects.strProductName + " deal details, error msg - " + strTxt;
					pageObjects.strGVStepName = "Failed to save the " + pageObjects.strProductName + " deal details, error msg - " + strTxt;
					fnScreenCapture(wDriver, strScreenPath.concat("dpmDealBookFailure") + "-"+ strDataPointer);
					System.out.println("[ERROR]   Failed to save the " + pageObjects.strProductName + " deal details, error msg - " + strTxt); 	} }
			else
			{ 	pageObjects.strScnStatus = false;
				pageObjects.strGVErrMsg = "Failed to find the Confirm Action button/screen";
				pageObjects.strGVStepName = "Failed to find the Confirm Action button/screen";
				fnScreenCapture(wDriver, strScreenPath.concat("dpmDealConfirmActionFailure") + "-"+ strDataPointer);
				System.out.println("[ERROR]   Failed to find the Confirm Action button/screen"); }	}
		else
		{	pageObjects.strScnStatus = false;
			pageObjects.strGVErrMsg = "Failed to find/ perform click on the SaveAs button";
			pageObjects.strGVStepName = "Failed to find/ perform click on the SaveAs button";
			fnScreenCapture(wDriver, strScreenPath.concat("dpmDealSaveAsFailure") + "-"+ strDataPointer);
			System.out.println("[ERROR]   Failed to find/ perform click on the SaveAs button");	}	}
//=================================================================================================================================================

	//Function to Book the SaveAs deal and get the new deal id, update in the results excel sheet
	public void fnBookDealSavedTemplate(String strDataSheet, String strDataPointer) throws Throwable {
		String strTxt = "", strVal = "", strTxt1="";
		boolean strSTxt = false;
		Robot robot = new Robot();
		int strIntTemp;
		if (fnObjectExists(objBtnBook, 2)!=0) {
			pageObjects.strGVErrMsg = "Successfully opened the Saved Deal screen";
			pageObjects.strGVStepName = "Successfully opened the Saved Deal screen";
			fnScreenCapture(wDriver, strScreenPath.concat("dpmDealBookOpenSuccess") + "-"+ strDataPointer);
			System.out.println("[INFO]   Successfully opened the Saved Deal screen ..............");
			
			if (pageObjects.strDealAmend == false) {
				fnObjectClick(objBtnBook); 
				robot.delay(1000);
				if (fnObjectExists(objConfirmAction, 2) != 0) {
					fnObjectClick(objConfirmAction);
					robot.delay(1000); }
				else if (fnObjectExists(objSaveAsConfirmAction, 2) != 0) {
					fnObjectClick(objSaveAsConfirmAction);
					robot.delay(1000); } }
			else if (pageObjects.strDealAmend == true) {
				fnObjectClick(objDealContinueAmendBtn); 
				robot.delay(1000);
				if (fnObjectExists(objConfirmAmend, 2) != 0) {
					fnObjectClick(objConfirmAmend);
					robot.delay(1000); }	}
			
			int strFound = 0;
			if (pageObjects.strScnStatus == true) {
				for(int intTemp = 0; intTemp<20; intTemp++) {
					strFound = -1;
					if ((fnObjectExists(objIDSuccessText, 2) != 0) || (fnObjectExists(objIDInBookingText, 2) != 0))
					{ 	
						if (fnObjectExists(objIDSuccessText, 2) != 0){
							Point objCoordinates = wDriver.findElement(By.xpath(objIDSuccessText)).getLocation();
							robot = new Robot();
							robot.mouseMove(objCoordinates.getX()+66, objCoordinates.getY()+86);
							strTxt = wDriver.findElement(By.xpath(objIDContent)).getText();
							
							strSTxt = (wDriver.findElement(By.xpath(objIDSuccessText)).getText().toLowerCase().contains("success"));
							robot.mouseMove(objCoordinates.getX()+65, objCoordinates.getY()+84);
							
							strTxt1 = wDriver.findElement(By.xpath(objIDContent)).getText(); }
						else if (fnObjectExists(objIDInBookingText, 2) != 0) {
							Point objCoordinates = wDriver.findElement(By.xpath(objIDInBookingText)).getLocation();
							robot = new Robot();
							robot.mouseMove(objCoordinates.getX()+66, objCoordinates.getY()+86);
							strTxt = wDriver.findElement(By.xpath(objIDInBookingContent)).getText();
							
							strSTxt = (wDriver.findElement(By.xpath(objIDInBookingText)).getText().toLowerCase().contains("in booking"));
							robot.mouseMove(objCoordinates.getX()+65, objCoordinates.getY()+84);
							
							strTxt1 = wDriver.findElement(By.xpath(objIDInBookingContent)).getText(); }
						if (strTxt1.trim().toLowerCase().contains("object") == true) {
							strTxt = "[Object Object]";
							strFound = 0; }
						else {
							fnUpdateExcelCellData(strDataSheetPath, strDataSheet, strDataPointer, "DPM_Deal_ID", strTxt);
							fnAppSync(pageObjects.intUpdateWaitTime);
							strFound = 1; } }
					else if (fnObjectExists(objIDErrText, 2) != 0)
						{	Point objCoordinates = wDriver.findElement(By.xpath(objIDErrText)).getLocation();
						robot = new Robot();
						robot.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+87);
						strTxt = wDriver.findElement(By.xpath(objServerErrorContent)).getText();
						strSTxt = (wDriver.findElement(By.xpath(objServerErrorContent)).getText().toLowerCase().contains("success"));
						robot.mouseMove(objCoordinates.getX()+64, objCoordinates.getY()+83);
						strFound = 0; pageObjects.strScnStatus = false;}
					else if (fnObjectExists(objFormErr, 1) != 0)
					{	strFound = 0; pageObjects.strScnStatus = false;
						strTxt = fnGetObjectAttributeValue(objFormErr, "outerText"); }
					robot.delay(1000);
					if (strFound != -1) break;} }

			if (strFound == 1) {
				if (strSTxt == true) {
					if (pageObjects.strDealAmend == false) {
						pageObjects.strGVStepName = "New Deal Booked Successfull and Deal Id - " + strTxt;
						System.out.println("[INFO]   New Deal Booked Successfull and Deal Id - " + strTxt);
						pageObjects.strDealBokID = strTxt;
					} else if (pageObjects.strDealAmend == true) {
						pageObjects.strGVStepName = "Deal Amendment Successfull and Deal Id - " + strTxt;
						System.out.println("[INFO]   Deal Amendment Successfull and Deal Id - " + strTxt);
						pageObjects.strDealBokID = strTxt; }
					pageObjects.strGVDealID = strTxt;
					fnScreenCapture(wDriver, strScreenPath.concat("dpmDealBookSuccess") + "-"+ strDataPointer);
					fnObjectClick(objPXVLabel);
					Point objCoordinates = wDriver.findElement(By.xpath(objIDSuccessText)).getLocation();
					robot = new Robot();
					robot.mouseMove(objCoordinates.getX()+76, objCoordinates.getY()+86);
					robot.mouseMove(objCoordinates.getX()+71, objCoordinates.getY()+81);
					wDriver.findElement(By.xpath(objIDSuccessText)).click();
					
					//step to update the deal id into the excel sheet
					fnAppSync(4);
					String strTime = fnGetCurrentDate("yyyy/mm/dd") + " " + fnGetCurrentTime();
					fnUpdateExcelCellData(strDataSheetPath, strDataSheet, strDataPointer, "Deal_DateTime", strTime);
					fnAppSync(pageObjects.intUpdateWaitTime); }
				else
					strFound = 0; }
			
			if ((strFound == 0) || (strFound == -1) || (pageObjects.strScnStatus == false)) {
				pageObjects.strScnStatus = false;
				if (pageObjects.strDealAmend == false) {
					pageObjects.strGVErrMsg = "Failed to Book the " + pageObjects.strProductName + " deal, error msg - " + strTxt;
					pageObjects.strGVStepName = "Failed to Book the " + pageObjects.strProductName + " deal, error msg - " + strTxt;
					fnScreenCapture(wDriver, strScreenPath.concat("dpmDealBookFailure") + "-"+ strDataPointer);
					System.out.println("[ERROR]   Failed to Book the " + pageObjects.strProductName + " deal, error msg - " + strTxt); }
				else if (pageObjects.strDealAmend == true) {
					pageObjects.strGVErrMsg = "Failed to Amend the " + pageObjects.strProductName + " deal, error msg - " + strTxt;
					pageObjects.strGVStepName = "Failed to Amend the " + pageObjects.strProductName + " deal, error msg - " + strTxt;
					fnScreenCapture(wDriver, strScreenPath.concat("dpmDealBookFailure") + "-"+ strDataPointer);
					System.out.println("[ERROR]   Failed to Amend the " + pageObjects.strProductName + " deal, error msg - " + strTxt); } }	} }
	
	//Function to search for the string in the array lists
	public int fnStringListCheck(String str, ArrayList <String> compareList) {
	    int intFound = -1;
	    for (int intTemp = 0; intTemp < compareList.size(); intTemp++){
	    	if (compareList.get(intTemp).trim().toLowerCase().contains(str.trim().toLowerCase()))
	    		intFound = intTemp;
	    	if (intFound!=-1)
	    		break; }
	    return intFound; }

	// Function to enter the values in the input tag field
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
			System.out.println("press under score");
			robot.keyPress(KeyEvent.VK_UNDERSCORE);
			robot.keyRelease(KeyEvent.VK_UNDERSCORE);
			break; } }

	/// Function to Read the value from the application & update the same in data sheet.
	public void fnReadUpdateDefaultValues(String strDataSheetPath, String dataSheet, String dataPointer,
			String strFieldName, String strColname, String strObject, String strAttribute) {
		
		String strVal=null;
		try {
			if (fnObjectExists(strObject, 2) != 0) {
				if (fnGetObjectAttributeValue(strObject, "nodeName").trim().toLowerCase()
						.contentEquals("select")) {
					WebElement elt = wDriver.findElement(By.xpath(strObject));
					strVal = new Select(elt).getFirstSelectedOption().getText();
					System.out.println("Fetching value from select tag> "+strFieldName); }
				else{
					strVal = fnGetObjectAttributeValue(strObject, strAttribute);
					if (!strVal.isEmpty()) {
						strVal = fnGetObjectAttributeValue(strObject, strAttribute);
						System.out.println("Reading value from field> "+strFieldName); }
					else{
						fnObjectClick(strObject);
						strVal = fnGetObjectAttributeValue(strObject, strAttribute);
						System.out.println("Reading value from field after click> "+strFieldName); } }
				if (!strVal.isEmpty()) {
					System.out.println("Value of " + strFieldName + " is: " + strVal);
					fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, strColname, strVal);
				} else 
					System.out.println("Value of " + strFieldName + " is empty!!"); }
			else{
				System.out.println("Object is not exsist for the xpath >> "+strObject);
				pageObjects.strCompareVal = false;}
		} catch (Throwable e) {
			pageObjects.strCompareVal = false;
			e.printStackTrace(); } }
	
	/// Function to check the stuats of a checkbox in application & update the status in data sheet.
	public void fnReadChkboxStatusUpdateDefaultValues(String strDataSheetPath, String dataSheet, String dataPointer,
			String strFieldName, String strColname, String strObject) {
		try {
			if (fnObjectExists(strObject, 2) != 0) {
				Boolean strVal = wDriver.findElement(By.xpath(strObject)).isSelected();
				if (strVal) {
					System.out.println("Checkbox " + strFieldName + " is selected!!");
					fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, strColname, "Checked");
				} else {
					System.out.println("Checkbox " + strFieldName + " is selected!!"); } }
		} catch (Throwable e) {
			pageObjects.strCompareVal = false;
			e.printStackTrace(); } }

	//Function for Credit Check ========================================================================================
	public void fnCreditCheck(String dataSheet, String dataPointer) throws Throwable {	
		String strVal = "";
		String strTxt = "";
		System.out.println("[INFO]    PXV - Starting Credit Check");		
		if (pageObjects.strScnStatus == true) {
			if(!(dataSheet.trim().toLowerCase().contains("eqo"))) {
				//Click on "Calc" button and then "Credit Check" to do credit check
				if(fnObjectExists(objCalcButton, 1) != 0) {
					wDriver.findElement(By.xpath(objCalcButton)).click();
					fnAppSync(2);
					if(fnObjectExists(objCreditCheck, 1) != 0)
						wDriver.findElement(By.xpath(objCreditCheck)).click(); } }
			else {
				if(fnObjectExists(objCreditCheckBtn, 1) != 0)
					wDriver.findElement(By.xpath(objCreditCheckBtn)).click();	}

			int strFound = 0;
			boolean strSTxt = false;
			for(int intTemp = 0; intTemp<20; intTemp++) {
				strFound = -1;
				if ((fnObjectExists(objIDSuccessText, 2) != 0) || (fnObjectExists(objIDSuccessText, 2) != 0))
				{ 	String strTxt1="";
					if (fnObjectExists(objIDSuccessText, 2) != 0){
						Point objCoordinates = wDriver.findElement(By.xpath(objIDSuccessText)).getLocation();
						Robot robot = new Robot();
						robot.mouseMove(objCoordinates.getX()+66, objCoordinates.getY()+86);
						strTxt = wDriver.findElement(By.xpath(objIDContent)).getText();
						strSTxt = (wDriver.findElement(By.xpath(objIDSuccessText)).getText().toLowerCase().contains("success"));
						robot.mouseMove(objCoordinates.getX()+65, objCoordinates.getY()+84);
						strTxt1 = wDriver.findElement(By.xpath(objIDContent)).getText();}
					else if (fnObjectExists(objIDInBookingText, 2) != 0){
						Point objCoordinates = wDriver.findElement(By.xpath(objIDInBookingText)).getLocation();
						Robot robot = new Robot();
						robot.mouseMove(objCoordinates.getX()+66, objCoordinates.getY()+86);
						strTxt = wDriver.findElement(By.xpath(objIDInBookingContent)).getText();
						strSTxt = (wDriver.findElement(By.xpath(objIDInBookingText)).getText().toLowerCase().contains("in booking"));
						robot.mouseMove(objCoordinates.getX()+65, objCoordinates.getY()+84);
						strTxt1 = wDriver.findElement(By.xpath(objIDInBookingContent)).getText();}
					
					if (strTxt1.trim().toLowerCase().contains("object") == true) {
						strTxt = "[Object Object]";
						strFound = 0; }
					else 
						strFound = 1; }
				else if (fnObjectExists(objIDErrText, 2) != 0)
				{	Point objCoordinates = wDriver.findElement(By.xpath(objIDErrText)).getLocation();
					Robot robot = new Robot();
					robot.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+87);
					strTxt = wDriver.findElement(By.xpath(objServerErrorContent)).getText();
					strSTxt = (wDriver.findElement(By.xpath(objServerErrorContent)).getText().toLowerCase().contains("success"));
					robot.mouseMove(objCoordinates.getX()+64, objCoordinates.getY()+83);
					strFound = 0; pageObjects.strScnStatus = false;}
					else if (fnObjectExists(objFormErr, 1) != 0) {
						strFound = 0; pageObjects.strScnStatus = false;
						strTxt = fnGetObjectAttributeValue(objFormErr, "outerText"); }
						fnAppSync(1);
						if (strFound != -1) break; }

			if (strSTxt == true) {				
				pageObjects.strGVStepName = "Credit Check is Successful";
				fnAppSync(1);
				fnScreenCapture(wDriver, strScreenPath.concat("dpmCreditCheckSuccess") + "-"+ dataPointer);
				fnObjectClick(objPXVLabel);
				if (fnObjectExists(objIDSuccessText, 2)!=1) {
					fnObjectClick(objIDSuccessText);
					fnObjectClick(objIDSuccessText); }
				fnAppSync(1);
				System.out.println("[INFO]   ######## Credit Check is Successful"); }				
			else {
				pageObjects.strScnStatus = false;
				pageObjects.strGVErrMsg = "Credit Check Failed";
				pageObjects.strGVStepName = "Credit Check Failed";
				fnAppSync(1);
				fnScreenCapture(wDriver, strScreenPath.concat("dpmCreditCheckFailed") + "-"+ dataPointer);
				System.out.println("[INFO]   ######## Credit Check Failed"); } }

		fnMouseMove(objCancelBtn);
		if(!(dataSheet.trim().toLowerCase().contains("eqo")))
		{
			if (pageObjects.strScnStatus == true) {			
				//Click on PV Calculator to calculate PV value
				if(fnObjectExists(objPVCalculator, 1) != 0) {
					fnAppSync(5);
					wDriver.findElement(By.xpath(objPVCalculator)).click(); }

				int strFound = 0;
				boolean strSTxt = false;
				for(int intTemp = 0; intTemp<20; intTemp++) {
					strFound = -1;
					if ((fnObjectExists(objIDSuccessText, 2) != 0) || (fnObjectExists(objIDInBookingText, 2) != 0))
					{ 	String strTxt1="";
						if (fnObjectExists(objIDSuccessText, 2) != 0){
							Point objCoordinates = wDriver.findElement(By.xpath(objIDSuccessText)).getLocation();
							Robot robot = new Robot();
							robot.mouseMove(objCoordinates.getX()+66, objCoordinates.getY()+86);
							strTxt = wDriver.findElement(By.xpath(objIDContent)).getText();
							strSTxt = (wDriver.findElement(By.xpath(objIDSuccessText)).getText().toLowerCase().contains("success"));
							robot.mouseMove(objCoordinates.getX()+65, objCoordinates.getY()+84);
		
							strTxt1 = wDriver.findElement(By.xpath(objIDContent)).getText(); }
						else if (fnObjectExists(objIDInBookingText, 2) != 0){
							Point objCoordinates = wDriver.findElement(By.xpath(objIDInBookingText)).getLocation();
							Robot robot = new Robot();
							robot.mouseMove(objCoordinates.getX()+66, objCoordinates.getY()+86);
							strTxt = wDriver.findElement(By.xpath(objIDInBookingContent)).getText();
							strSTxt = (wDriver.findElement(By.xpath(objIDInBookingText)).getText().toLowerCase().contains("in booking"));
							robot.mouseMove(objCoordinates.getX()+65, objCoordinates.getY()+84);
		
							strTxt1 = wDriver.findElement(By.xpath(objIDInBookingContent)).getText(); }
						if (strTxt1.trim().toLowerCase().contains("object") == true) {
							strTxt = "[Object Object]";
							strFound = 0; }
						else 
							strFound = 1; }
					else if (fnObjectExists(objIDErrText, 2) != 0) {
						Point objCoordinates = wDriver.findElement(By.xpath(objIDErrText)).getLocation();
						Robot robot = new Robot();
						robot.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+87);
						strTxt = wDriver.findElement(By.xpath(objServerErrorContent)).getText();
						strSTxt = (wDriver.findElement(By.xpath(objServerErrorContent)).getText().toLowerCase().contains("success"));
						robot.mouseMove(objCoordinates.getX()+64, objCoordinates.getY()+83);
						strFound = 0; pageObjects.strScnStatus = false;}
						else if (fnObjectExists(objFormErr, 1) != 0)
						{	strFound = 0; pageObjects.strScnStatus = false;
						strTxt = fnGetObjectAttributeValue(objFormErr, "outerText"); }
						fnAppSync(1);
						if (strFound != -1) break; }

				if (strSTxt == true) {	
					fnAppSync(1);
					pageObjects.strGVStepName = "PV calculation is Completed";
					fnScreenCapture(wDriver, strScreenPath.concat("dpmPVCalculationComplete") + "-"+ dataPointer);
					fnObjectClick(objPXVLabel);
					if (fnObjectExists(objIDSuccessText, 2)!=1) {
						fnObjectClick(objIDSuccessText);
						fnObjectClick(objIDSuccessText); }
					fnAppSync(1);
					String strPVValue = wDriver.findElement(By.xpath(objPVCalcValue)).getText().trim();
					fnUpdateExcelCellData(strDataSheetPath, dataSheet, dataPointer, "PV", strPVValue);
					System.out.println("[INFO]   ######## PV calculation is Completed"); }				
				else {	
					fnAppSync(1);
					pageObjects.strGVErrMsg = "PV calculation Failed";
					pageObjects.strGVStepName = "PV calculation Failed";					
					fnScreenCapture(wDriver, strScreenPath.concat("dpmPVCalculationFailed") + "-"+ dataPointer);
					System.out.println("[INFO]   ######## PV calculation Failed"); } }
			fnMouseMove(objCancelBtn);
			fnAppSync(5); }		

		//Step to read and update Credit check values in test data sheet
		if (pageObjects.strScnStatus == true) {
			if (fnObjectExists(objCreditCheckValuesTable, 2) != 0) {
				List<WebElement> totalRows = wDriver.findElements(By.xpath(objCreditCheckTableTotalRows));  
				for (int row = 1; row <= totalRows.size(); row++) 
				{	
					String objTotolColsPerRow = objCreditCheckTableTotalColsPerRow.replace("ROW", String.valueOf(row));
					List<WebElement> totalCols = wDriver.findElements(By.xpath(objTotolColsPerRow));
					String strField = "";
					for (int col = 1; col <= totalCols.size(); col++) {						
						String objColValue = objCreditCheckTableColValue.replace("ROW", String.valueOf(row)).replace("COL", String.valueOf(col));
						if(row == 1 && col == 2) {
							strVal =  wDriver.findElement(By.xpath(objColValue)).getAttribute("value").trim();
							if(strVal.isEmpty() == false) {
								fnUpdateExcelCellData(strDataSheetPath, dataSheet, dataPointer, "CC_DealID" , strVal);
								System.out.println("[INFO]   == == == Successfully written, Field - "+ "\"" + "CC_DealID" + "\"" + "   value - " + "\"" + strVal + "\"" + " in Test Data file."); } }
						else 
						{
							String strFieldName = "";
							if(col == 1) 
								strField =  wDriver.findElement(By.xpath(objColValue)).getAttribute("value").trim();								
							else if(col == 2) {								
								strVal =  wDriver.findElement(By.xpath(objColValue)).getAttribute("value").trim();
								if(strVal.isEmpty() == false) {
									strFieldName = "CC_" + strField + "1" ;
									fnUpdateExcelCellData(strDataSheetPath, dataSheet, dataPointer, strFieldName , strVal);
									System.out.println("[INFO]   == == == Successfully written, Field - "+ "\"" + strFieldName + "\"" + "   value - " + "\"" + strVal + "\"" + " in Test Data file.");
								}
							}
							else {
								strVal =  wDriver.findElement(By.xpath(objColValue)).getAttribute("value").trim();
								if(strVal.isEmpty() == false) {
									strFieldName = "CC_" + strField + "2" ;
									fnUpdateExcelCellData(strDataSheetPath, dataSheet, dataPointer, strFieldName , strVal);
									System.out.println("[INFO]   == == == Successfully written, Field - "+ "\"" + strFieldName + "\"" + "   value - " + "\"" + strVal + "\"" + " in Test Data file.");
								}
							}
						}
					}
				}
			}
		}
		System.out.println("[INFO]   ######## Credit Check completed."); }
	
//Function to set the data in the CASH FLOW tab ============================================================================================
	public void fnSetCashFlowTabData(String strDataSheet, String strDataPointer) throws Throwable {
		int strIntTemp;
		String strVal="";
		Robot robot = new Robot();
		if (fnObjectExists(objCashFlowBtn, 3)!=0){
			
			//Steps to enter the values in CASH FLOW Tab - FIXED PAY SIDE
			strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "CashFlow_FixedPay");
			if (strVal.trim().length()>0)
				if (fnObjectExists(objFPDropdown, 1)!=0)
					strIntTemp = fnSelectValue(objFPDropdown, strVal);
			
			strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "CashFlow_FPAdjType");
			if (strVal.trim().length()>0)
				if (fnObjectExists(objFPAdjType, 1)!=0)
					strIntTemp = fnSelectValue(objFPAdjType, strVal);

			strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "CashFlow_FPFreq");
			if (strVal.trim().length()>0)
				if (fnObjectExists(objFPFreq, 1)!=0)	{
					wDriver.findElement(By.xpath(objFPFreq)).clear();
					wDriver.findElement(By.xpath(objFPFreq)).click();
					fnAppSync(2);
					wDriver.findElement(By.xpath(objFPFreq)).sendKeys(strVal + Keys.ENTER); }
			
			strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "CashFlow_FPUnit");
			if (strVal.trim().length()>0)
				if (fnObjectExists(objFPUnit, 1)!=0) {
					strIntTemp = fnSelectValue(objFPUnit, strVal);
					wDriver.findElement(By.xpath(objFPUnit)).sendKeys(Keys.TAB); }
			
			strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "CashFlow_FPAmount").trim();
			if (strVal.trim().length()>0) {
				try{
					wDriver.findElement(By.xpath(objFPAmount)).click();
					robot.delay(100);
					wDriver.findElement(By.xpath(objFPAmount)).click();
					robot.delay(100); }
				catch (Exception exp) {
					try { 
						wDriver.findElement(By.xpath(objFPAmount)).click();
						robot.delay(100); }
					catch (Exception exp1){
						System.out.println("[INFO]    trying to search for the input field"); }	}
				try{
					wDriver.findElement(By.xpath(objFPAmtInput)).sendKeys(strVal.trim()); }
				catch (Exception exp){
					try{
						wDriver.findElement(By.xpath(objFPAmtInput)).sendKeys(strVal.trim()); }
					catch (Exception exp1){
						System.out.println("[INFO]    trying to enter details for the input field"); }	} }

			//Steps to enter the values in CASH FLOW Tab - FLOAT RECEIVE SIDE
			strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "CashFlow_FloatReceive");
			if (strVal.trim().length()>0)
				if (fnObjectExists(objFRDropdown, 1)!=0)
					strIntTemp = fnSelectValue(objFRDropdown, strVal);
			
			strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "CashFlow_FRAdjType");
			if (strVal.trim().length()>0)
				if (fnObjectExists(objFRAdjType, 1)!=0)
					strIntTemp = fnSelectValue(objFRAdjType, strVal);
			
			strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "CashFlow_FRFreq");
			if (strVal.trim().length()>0) 
				if (fnObjectExists(objFRFreq, 1)!=0)	{
					wDriver.findElement(By.xpath(objFRFreq)).clear();
					wDriver.findElement(By.xpath(objFRFreq)).click();
					fnAppSync(1);
					wDriver.findElement(By.xpath(objFRFreq)).sendKeys(strVal + Keys.ENTER); }
			
			strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "CashFlow_FRUnit");
			if (strVal.trim().length()>0)
				if (fnObjectExists(objFRUnit, 1)!=0)
					strIntTemp = fnSelectValue(objFRUnit, strVal);
			
			strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "CashFlow_FRAmount");
			if (strVal.trim().length()>0) 
				if (fnObjectExists(objFRAmount, 1)!=0) {
					wDriver.findElement(By.xpath(objFRAmount)).click();
					fnAppSync(1);
					wDriver.findElement(By.xpath(objFRAmtInput)).sendKeys(strVal + Keys.ENTER); }

			pageObjects.strGVStepName = "Successfully entered the data in the Cash Flow tab ";
			fnScreenCapture(wDriver, strScreenPath.concat("dpmBasisCashFlow") + "-"+ strDataPointer); } }

//Function to read PV value from QUICK ENTRY and update in excel sheet ============================================================================================
	public void fnReadPVCalcQuckEntryScreen(String strDataSheet, String strDataPointer) throws Throwable {
		
		if (fnObjectExists(objPVCalQuickEntry, 2)!=0){
			fnObjectClick(objPVCalQuickEntry);
			
			int strFound = -1;
			String strTxt="";
			boolean strSTxt = false;
			for(int intTemp = 0; intTemp<20; intTemp++) {
				strFound = -1;
				if (fnObjectExists(objIDSuccessText, 1) != 0)
				{ 	Point objCoordinates = wDriver.findElement(By.xpath(objIDSuccessText)).getLocation();
					Robot robot = new Robot();
					robot.mouseMove(objCoordinates.getX()+66, objCoordinates.getY()+86);
					strTxt = wDriver.findElement(By.xpath(objIDContent)).getText();
					strSTxt = (wDriver.findElement(By.xpath(objIDSuccessText)).getText().toLowerCase().contains("success"));
					robot.mouseMove(objCoordinates.getX()+65, objCoordinates.getY()+84);
					String strTxt1 = wDriver.findElement(By.xpath(objIDContent)).getText();
					fnAppSync(1);
					robot.mouseMove(objCoordinates.getX()+70, objCoordinates.getY()+75);
					if (fnObjectExists(objIDSuccessText, 2)!=1) {
						fnObjectClick(objIDSuccessText);
						fnObjectClick(objIDSuccessText); }
					fnAppSync(1);
					if (strTxt1.trim().toLowerCase().contains("object") == true)
						strFound = 0;
					else 
						strFound = 1; }
				else if (fnObjectExists(objIDErrText, 1) != 0)
				{	Point objCoordinates = wDriver.findElement(By.xpath(objIDErrText)).getLocation();
					Robot robot = new Robot();
					robot.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+87);
					strTxt = wDriver.findElement(By.xpath(objServerErrorContent)).getText();
					strSTxt = (wDriver.findElement(By.xpath(objServerErrorContent)).getText().toLowerCase().contains("success"));
					robot.mouseMove(objCoordinates.getX()+64, objCoordinates.getY()+83);
					strFound = 0; pageObjects.strScnStatus = false; }
				else if (fnObjectExists(objFormErr, 1) != 0)
				{	strFound = 0;
					strTxt = fnGetObjectAttributeValue(objFormErr, "outerText");
					pageObjects.strScnStatus = false;}
					fnAppSync(1);
				if (strFound != -1) break;}
			
			if (strFound == 1){
				if (fnObjectExists(objPVCalcValueQuickEntry, 1) != 0){
					String strVal = fnGetObjectAttributeValue(objPVCalcValueQuickEntry, "outerText");
					if (strVal.trim().length()==0) strVal = "NA";
					fnUpdateExcelCellData(strDataSheetPath, strDataSheet, strDataPointer, "PV_VALUE", strVal); 
					System.out.println("[INFO]    Successfully retrieved the PV VALUE from Quick Entry Form - " + strVal); }
				else {
					fnUpdateExcelCellData(strDataSheetPath, strDataSheet, strDataPointer, "PV_VALUE", "NA");
					System.out.println("[WARNING]    Unable to retrieve the PV VALUE from Quick Entry Screen"); } } 
			else{
				System.out.println("[WARNING]    Unable to retrieve the PV VALUE from Quick Entry Screen");
				fnUpdateExcelCellData(strDataSheetPath, strDataSheet, strDataPointer, "PV_VALUE", "NA"); } }
		else {
			System.out.println("[WARNING]    Unable to find the PV CALC button in the Quick Entry Form"); } 
		
	}

	//Function to Verify the POST PV value from QUICK ENTRY and update in excel sheet ============================================================================================
	public void fnVerifyPostPVCalcQuckEntryScreen(String strDataSheet, String strDataPointer) throws Throwable {
		
		String strPVValue = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, "PV_VALUE");
		if (strPVValue.trim().length()>0)
			if (strPVValue.trim().toLowerCase().contentEquals("NA")==false)
				if (fnObjectExists(objPVCalQuickEntry, 2)!=0){
					fnObjectClick(objPVCalQuickEntry);
					
					int strFound = -1;
					String strTxt="";
					boolean strSTxt = false;
					for(int intTemp = 0; intTemp<20; intTemp++) {
						strFound = -1;
						if (fnObjectExists(objIDSuccessText, 1) != 0)
						{ 	Point objCoordinates = wDriver.findElement(By.xpath(objIDSuccessText)).getLocation();
							Robot robot = new Robot();
							robot.mouseMove(objCoordinates.getX()+66, objCoordinates.getY()+86);
							strTxt = wDriver.findElement(By.xpath(objIDContent)).getText();
							strSTxt = (wDriver.findElement(By.xpath(objIDSuccessText)).getText().toLowerCase().contains("success"));
							robot.mouseMove(objCoordinates.getX()+65, objCoordinates.getY()+84);
							String strTxt1 = wDriver.findElement(By.xpath(objIDContent)).getText();
							fnAppSync(1);
							if (fnObjectExists(objIDSuccessText, 2)!=1) {
								fnObjectClick(objIDSuccessText);
								fnObjectClick(objIDSuccessText); }
							fnAppSync(1);
							if (strTxt1.trim().toLowerCase().contains("object") == true)
								strFound = 0;
							else 
								strFound = 1; }
						else if (fnObjectExists(objIDErrText, 1) != 0)
						{	Point objCoordinates = wDriver.findElement(By.xpath(objIDErrText)).getLocation();
							Robot robot = new Robot();
							robot.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+87);
							strTxt = wDriver.findElement(By.xpath(objServerErrorContent)).getText();
							strSTxt = (wDriver.findElement(By.xpath(objServerErrorContent)).getText().toLowerCase().contains("success"));
							robot.mouseMove(objCoordinates.getX()+64, objCoordinates.getY()+83);
							strFound = 0; pageObjects.strScnStatus = false; }
						else if (fnObjectExists(objFormErr, 1) != 0)
						{	strFound = 0;
							strTxt = fnGetObjectAttributeValue(objFormErr, "outerText");
							pageObjects.strScnStatus = false;}
							fnAppSync(1);
						if (strFound != -1) break;}
					
					if (strFound == 1){
						if (fnObjectExists(objPVCalcValueQuickEntry, 1) != 0){
							String strVal = fnGetObjectAttributeValue(objPVCalcValueQuickEntry, "outerText");
							//step to verify the PV value pre and post
							fnNumericObjCompareValues(strDataSheetPath, strDataSheet, strDataPointer, "PV_VALUE", "PV_VALUE", objPVCalcValueQuickEntry, "outerText");
							System.out.println("[INFO]    Successfully retrieved the PV VALUE from Quick Entry Form - " + strVal); 
							pageObjects.strCompareVal = true; }
						else {
							System.out.println("[WARNING]    Unable to retrieve/verify the PV VALUE from Quick Entry Screen"); } } 
					else{
						System.out.println("[WARNING]    Unable to retrieve/verify the PV VALUE from Quick Entry Screen"); } }
				else {
					System.out.println("[WARNING]    Unable to find the PV CALC button in the Quick Entry Form"); }
	
	}
	
	//Function to read deal default values and write in excel sheet
	public void fnReadWriteDefaultValues(String strDataSheetPath,String dataSheet,String dataPointer, String strFieldName, String strColname, String strObject,String strAttribute) throws Throwable {

		String strActVal = null;
		try {
			if (fnObjectExists(strObject, 1) != 0) {
				//step to read the object attribute property value
				if (fnGetObjectAttributeValue(strObject, "nodeName").trim().toLowerCase().contentEquals("select")) 	{
					WebElement elt = wDriver.findElement(By.xpath(strObject));
					strActVal = new Select(elt).getFirstSelectedOption().getText().trim();  
					System.out.println("[INFO]   Field "+ "\"" + strFieldName + "\"" + " value is >>>>>>>> " + strActVal); }
				else {
					strActVal = fnGetObjectAttributeValue(strObject, strAttribute).trim();
					System.out.println("[INFO]   Field "+ "\"" + strFieldName + "\"" + " value is >>>>>>>> " + strActVal); }

				//step to check if its NOT empty or NULL and then write in test data file
				if(!(strActVal.isEmpty())) {
					fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, strColname, strActVal);
					//System.out.println("[INFO]    == == == Successfully written - " + "\"" + strFieldName + "\"" + " value :in Test Data file.");
					System.out.println("Field Name : " + strFieldName + " - Value : " + strActVal);
					}
				else
					System.out.println("[INFO]    ## ## ## Field " + "- \"" + strFieldName + "\"" + " is empty or NULL."); }
		}catch(Exception e) { pageObjects.strCompareVal = false; }	}
	
	//Function to clear the Notifications ====================================================================================
	public void fnClearNotifications() throws Throwable {
		try{
			if (fnObjectExists(objMrktOfflineErr, 2)!=0) {
				Point objCoordinates = wDriver.findElement(By.xpath(objMrktOfflineErr)).getLocation();
				Robot robot = new Robot();
				robot.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+87);
				fnAppSync(1);
				robot.mouseMove(objCoordinates.getX()+64, objCoordinates.getY()+83); 
				wDriver.findElement(By.xpath(objMrktOfflineErr)).click(); }
			else if (fnObjectExists(objMrktOfflineSuccess, 2)!=0) {
				Point objCoordinates = wDriver.findElement(By.xpath(objMrktOfflineSuccess)).getLocation();
				Robot robot = new Robot();
				robot.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+87);
				fnAppSync(1);
				robot.mouseMove(objCoordinates.getX()+64, objCoordinates.getY()+83); 
				wDriver.findElement(By.xpath(objMrktOfflineSuccess)).click(); }

			if (fnObjectExists(objNotificationBellIcon, 1) != 0) {
				try{
					fnObjectClick(objNotificationBellIcon);
					fnAppSync(2); }
				catch (Exception exp){
					System.out.println("[INFO]    Failed to click on the Notification icon");	}
				
				try{
					wDriver.findElement(By.xpath(objNotificationBellIcon)).click();
					fnAppSync(2); }
				catch (Exception exp){
					System.out.println("[INFO]    Failed to click on the Notification icon"); } 
				
				if (fnObjectExists(objPXVLabel, 1) != 0)
					fnObjectClick(objPXVLabel);
				fnAppSync(1);
				System.out.println("[INFO]   Successfully cleared the Notifications..... .....");

				//step to close the notifications window
				//if (fnObjectExists(objBtnCloseNotificationSideBar, 1) != 0)
				//	fnObjectClick(objBtnCloseNotificationSideBar); 
			} }
		catch(Exception exp)
		{ System.out.println("[ERROR]   Unable to click on the Notifications Bell Icon......"); }	}
	
	// function to get deal id from notifications window and compare with excel ==============================================
	public void fnVerifyDealIdInNotification(String dataSheet, String dataPointer) throws Throwable{
		String excDealId = fnGetExcelData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "DPM_Deal_ID");
		String notificationContent , notificationDealId ;
		Robot robot = new Robot();
		// click on notification bell 
		if(fnObjectExists(objNotificationBellIcon, 2) != 0){
			for (int intTemp=0;intTemp<25;intTemp++){
				try{
					robot.delay(8000);
					wDriver.findElement(By.xpath(objNotificationBellIcon)).click();
					intTemp = 50;	}
				catch (Exception exp){ 
					System.out.println("[INFO]    Trying to click on the object - " + objNotificationBellIcon + " = " + intTemp); } } }
		fnAppSync(1);
		
		String strNotObj1 = objDealNotificationSucc.replace("DEAL-ID", excDealId.trim().toUpperCase());
		String strNotObj2 = objDealNotificationFail.replace("DEAL-ID", excDealId.trim().toUpperCase());
		String strNotObj3 = objDealNotificationSuccMsg.replace("DEAL-ID", excDealId.trim().toUpperCase());
		String strNotObj4 = objDealNotificationFailMsg.replace("DEAL-ID", excDealId.trim().toUpperCase());

		pageObjects.strGVErrMsg = "Notification popup details for Deal Id - " + excDealId + " in the Notification Window";
		pageObjects.strGVStepName = "Notification popup details for Deal Id - " + excDealId + " in the Notification Window";
		fnScreenCapture(wDriver, strScreenPath.concat("DealIdVer1") + "-"+ pageObjects.strGVScrDataPointer);
		System.out.println("[INFO]   Verified Deal Id In Notifications Window - "+excDealId);
		
		if (fnObjectExists(objBtnDismissAll, 1) != 0)
			fnObjectClick(objBtnDismissAll);
		fnAppSync(1);	

		// close the notification window
		if(fnObjectExists(objBtnCloseNotificationSideBar, 2) != 0)
			for (int intTemp=0;intTemp<25;intTemp++) {
				try{
					robot.delay(8000);
					wDriver.findElement(By.xpath(objBtnCloseNotificationSideBar)).click();
					intTemp = 50;	}
				catch (Exception exp){ 
					System.out.println("[INFO]    Trying to click on the object - " + objBtnCloseNotificationSideBar + " = " + intTemp); } } }

	// function to select default workspace ==============================================================================================
	public void fnSelectDefaultWorkspace() throws Throwable{
		try{
			if(fnObjectExists(objWorkspaceButton, 2) != 0)
				fnObjectClick(objWorkspaceButton);
			fnAppSync(1);
			if(fnObjectExists(objDefaultWorkspaceLink, 2) != 0)
				fnObjectClick(objDefaultWorkspaceLink);
			pageObjects.strScnStatus = true;
			System.out.println("[INFO]   Successfully selected the Default Workspace......");
		}catch(Exception e){
			System.out.println("[ERROR]   Unable to select the Default Workspace button......"); } }
	
	
// function to select / create workspace =============================================================================================
	public void fnSelectWorkspace(String workspace)throws Throwable{
		if(fnObjectExists(objWorkspaceButton, 2) != 0){
			try{
				fnObjectClick(objWorkspaceButton1);
			}catch(Exception e){
				fnAppSync(3);
				Actions actions = new Actions(wDriver);
				actions.moveToElement(wDriver.findElement(By.xpath(objWorkspaceButton1))).click().build().perform();	}
			fnAppSync(2);
			String temp = objWorkspace;
			temp = temp.replace("WORKSPACE", workspace.trim());
			if(fnObjectExists(temp, 3) != 0){
				try{
					fnObjectClick(temp);
				}catch(Exception e){
					fnAppSync(3);
					Actions actions = new Actions(wDriver);
					actions.moveToElement(wDriver.findElement(By.xpath(temp))).click().build().perform();	}
				fnAppSync(2);
				//pageObjects.strGVStepName = "Successfully Selected The Workspace - "+workspace;
				System.out.println("[ERROR]		Successfully Selected The Workspace - "+workspace);
				//fnScreenCapture(wDriver, strScreenPath.concat("workspacenotfound"));
			}else {
				//pageObjects.strScnStatus = false;
				//pageObjects.strGVStepName = "Workspace Not Found, Creating The new Workspace - "+workspace;
				System.out.println("[ERROR]		Workspace Not Found, Creating The new Workspace - "+workspace);
				//fnScreenCapture(wDriver, strScreenPath.concat("workspacenotfound"));
				//pageObjects.strGVErrMsg = "Workspace Not Found, Creating The new Workspace - "+workspace;

				fnClearDPMDashboard();
				
				if(fnObjectExists(objMenuLnk, 2) != 0){
					fnObjectClick(objMenuLnk);
					fnAppSync(2);
					
					if(fnObjectExists(objNewWorkspace, 2) != 0){
						fnObjectClick(objNewWorkspace);
						fnAppSync(5);
						
						if(fnObjectExists(objNewWorkspaceText, 2) != 0){
							wDriver.findElement(By.xpath(objNewWorkspaceText)).click();
							fnAppSync(1);
							wDriver.findElement(By.xpath(objNewWorkspaceText)).clear();
							fnAppSync(1);
							wDriver.findElement(By.xpath(objNewWorkspaceText)).sendKeys(workspace);
							fnAppSync(1);
							
							pageObjects.strGVStepName = "Successfully Entered Workspace To create - "+workspace;
							System.out.println("[ERROR]		Successfully Entered Workspace To create - "+workspace);
							fnScreenCapture(wDriver, strScreenPath.concat("workspacenotfound"));
							
							if(fnObjectExists(objNewWorkspaceCreateButton, 2) != 0){
								try{
									fnObjectClick(objNewWorkspaceCreateButton);
								}catch(Exception e){
									fnAppSync(3);
									Actions actions = new Actions(wDriver);
									actions.moveToElement(wDriver.findElement(By.xpath(objNewWorkspaceCreateButton))).click().build().perform();	}
								
								//pageObjects.strGVStepName = "Successfully Created and selected workspace - "+workspace;
								System.out.println("[ERROR]		Successfully Created and selected workspace - "+workspace);
								//fnScreenCapture(wDriver, strScreenPath.concat("workspacenotfound"));
							}else{
								//pageObjects.strScnStatus = false;
								//pageObjects.strGVStepName = "Unable to find Create Button";
								System.out.println("[ERROR]		Unable to find Create Button");
								//fnScreenCapture(wDriver, strScreenPath.concat("workspacenotfound"));
								//pageObjects.strGVErrMsg = "Unable to find Create Button";
								}
						}else{
							pageObjects.strScnStatus = false;
							pageObjects.strGVStepName = "Unable to find The Text Field To Enter Workspace";
							System.out.println("[ERROR]		Unable to find The Text Field To Enter Workspace");
							fnScreenCapture(wDriver, strScreenPath.concat("workspacenotfound"));
							pageObjects.strGVErrMsg = "Unable to find The Text Field To Enter Workspace";}
					}else{
						pageObjects.strScnStatus = false;
						pageObjects.strGVStepName = "Unable to find New Workspace Option In The Menu Options";
						System.out.println("[ERROR]		Unable to find New Workspace Option In The Menu Options");
						fnScreenCapture(wDriver, strScreenPath.concat("workspacenotfound"));
						pageObjects.strGVErrMsg = "Unable to find New Workspace Option In The Menu Options";}
				}else{
					pageObjects.strScnStatus = false;
					pageObjects.strGVStepName = "Unable To Find The Menu Link";
					System.out.println("[ERROR]		Unable To Find The Menu Link");
					fnScreenCapture(wDriver, strScreenPath.concat("workspacenotfound"));
					pageObjects.strGVErrMsg = "Unable To Find The Menu Link";}}
		}else{
			pageObjects.strScnStatus = false;
			pageObjects.strGVStepName = "Unable to find Workspace Drop Down";
			System.out.println("[ERROR]		Unable to find Workspace Drop Down");
			fnScreenCapture(wDriver, strScreenPath.concat("workspacenotfound"));
			pageObjects.strGVErrMsg = "Unable to find Workspace Drop Down";}	}
	
//Function to read the DB values ==========================================================================
	public String fnReadDBValues(String strQuery, String strDB, int intFieldVal) throws Throwable{
		String strDBUrl="";
		String strRetValue = "";
		String strDBUser = "mtm_oper";
		String strDBpwd = "mtm_oper1";
		
		if ((strDB.trim().toLowerCase().contains("ist4")) || (strDB.trim().toLowerCase().contains("azure")))
			strDBUrl = "jdbc:sybase:Tds:172.22.141.115:4010";
		else if (strDB.trim().toLowerCase().contentEquals("uat"))
			strDBUrl = "jdbc:sybase:Tds:172.20.184.78:7900";
		
		try { Class.forName("com.sybase.jdbc.SybDriver"); }
	    catch (ClassNotFoundException cnfe) { }
		
		try { System.out.println("[INFO]    Opening a connection.");
			Connection con = DriverManager.getConnection(strDBUrl, strDBUser, strDBpwd);
			Statement querySet = con.createStatement();
			System.out.println("[INFO]    Successfully opened connection.");
			ResultSet rs = querySet.executeQuery(strQuery);
			System.out.println("[INFO]    Successfully executed the query. .. rs : " + rs.getFetchSize());
			int intTemp = 0;
			while(rs.next()) {
				intTemp = intTemp + 1;
				strRetValue = rs.getString(intFieldVal);
				System.out.println("[INFO]     >>>> " + rs.getString(intFieldVal)); }
			rs.close();
			querySet.close();
			con.close(); }
		catch (SQLException sqe) {
			System.out.println("[ERROR]    Failed to the DB connection.");
			System.out.println("[ERROR]    Unexpected exception : " +
				   sqe.toString() + ", sqlstate = " +
				   sqe.getSQLState()); }
		if (strRetValue.trim().length()==0) strRetValue = "NA";
		System.out.println("DB Return Value : " + strRetValue);
		return strRetValue; }

// function to update the command bar values in excel ============================================================================
	public void fnUpdateCmdBarValuesInExcel(String dataSheet, String dataPointer)throws Throwable{
		String strVal;
		try{
			// step to update trade date
			strVal = fnGetObjectAttributeValue(objTradeDate, "value");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "TradeDate", strVal);

			// counter party
			strVal = fnGetObjectAttributeValue(objCounteParty, "value");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "CounterParty", strVal);

			// initial strike
			strVal = fnGetObjectAttributeValue(objStartDate, "value");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "InitialStrike", strVal);

			// book
			strVal = fnGetObjectAttributeValue(objBookType, "value");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Book", strVal);

			// expiry
			strVal = fnGetObjectAttributeValue(objExpiryDate, "value");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Expiry", strVal);

			// Booking Point
			strVal = fnGetObjectAttributeValue(objBookingPoint, "value");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "BookingPoint", strVal);

			// deal revenue
			strVal = fnGetObjectAttributeValue(objDealRevenue, "outerText");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "DealRevenue", strVal);

			// marketer Id
			strVal = fnGetObjectAttributeValue(objMarketerID, "value");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "MarketerID", strVal);

			// groups
			strVal = fnGetObjectAttributeValue(objGroups, "outerText");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Groups", strVal);

			// national
			strVal = fnGetObjectAttributeValue(objNotional, "outerText");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Notional", strVal);

			// Units
			strVal = fnGetObjectAttributeValue(objUnits, "value");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Units", strVal);

			// Style
			strVal = fnGetObjectAttributeValue(objStyle, "value");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Style", strVal);

			// Underlying
			strVal = fnGetObjectAttributeValue(objUnderlying, "value");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Underlying", strVal);

			// spot type
			strVal = fnGetObjectAttributeValue(objSpotType, "value");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "SpotType", strVal);

			// spot price
			strVal = fnGetObjectAttributeValue(objSpanSpotPrice, "outerText");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "SpotPrice", strVal);
			
			// Strike %
			strVal = fnGetObjectAttributeValue(objStrikePercent, "value");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "StrikePercentage", strVal);
			
			// strike price
			strVal = fnGetObjectAttributeValue(objStrikePrice, "outerText");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "StrikePrice", strVal);

			// cap price
			strVal = fnGetObjectAttributeValue(objCapPrice, "outerText");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "CapPrice", strVal);
			
			// cap %
			strVal = fnGetObjectAttributeValue(objCapPercent, "value");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "CapPercentage", strVal);

			// premium
			strVal = fnGetObjectAttributeValue(objPremium, "outerText");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Premium", strVal);

			// premium settlement date
			strVal = fnGetObjectAttributeValue(objPremiumSettleDate, "value");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "PremiumSettlementDate", strVal);

			// settlement period
			strVal = fnGetObjectAttributeValue(objSettlementPeriod, "value");
			fnUpdateExcelCellData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "SettelementPeriod", strVal);

			System.out.println("[INFO]	Successfully Updated The Values In The Excel Sheet");

		}catch(Exception e){
			e.printStackTrace();
			System.out.println("[ERROR]	Exception Occured While Updating Commandbar Excel Data");	}	}


// function to enter the command bar details thru the command bar =============================================================================================================
	public void fnEnterCommandDealThroughCmdBar(String dataSheet, String dataPointer, String product, String strFunctionality) throws Throwable {
		String strTxt = new String();
		Robot robot = new Robot();
		int strFound = 0;
		
		//function call to clear the dash board
		fnClearDPMDashboard();

		//function call to clear the NOTIFICATIONS
		fnClearNotifications();

		String strCommandBarFlag = fnGetExcelData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "UseCommand_Bar");
		if(strCommandBarFlag.trim().toLowerCase().equals("y")) {
			// get the deal command
			//String dealCommand = fnCommLib.fnGetExcelData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "CommandCenter");
			String dealCommand = fnGetExcelData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Deal_Command");

			//TRADE, EXPIRY & SETTLEMENT - DATE TO RETRIVE FROM DATABASE - NEED TO FIND THE COLUMN & TABLE NAMES....			
			if (product.trim().toLowerCase().contentEquals("eqo")) {
				if (strFunctionality.trim().toLowerCase().contains("dealbook")) {
					if (dealCommand.trim().toLowerCase().contains("tradedate"))
						dealCommand = dealCommand.toLowerCase().replace("tradedate", pageObjects.strTradeDate);
			
					if (dealCommand.trim().toLowerCase().contains("expirydate"))
						dealCommand = dealCommand.toLowerCase().replace("expirydate", pageObjects.strExpiryDate);
			
					if (dealCommand.trim().toLowerCase().contains("settlementdate"))
						dealCommand = dealCommand.toLowerCase().replace("settlementdate", pageObjects.strSettlementDate); } }
			
			if (strFunctionality.trim().toLowerCase().contains("dealunwind")) {
					String strTemp0 = fnGetExcelData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "DPM_Deal_ID");
					String strTemp1 = fnGetExcelData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Unwind_Fee");
					String strTemp2 = fnGetExcelData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Unwind_FeeCurr");
					String strTemp3 = fnGetExcelData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Unwind_FeeSettlDate");
					String strTemp4 = fnGetExcelData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Unwind_FeeBook");
					dealCommand = "/unwind " + strTemp0 + " ";
					if (strTemp1.trim().length()>0){
						dealCommand = dealCommand + strTemp1 + " ";
						if (strTemp2.trim().length()>0){
							dealCommand = dealCommand + strTemp2 + " ";
							if (strTemp3.trim().length()>0){
								dealCommand = dealCommand + strTemp3 + " ";
								if (strTemp4.trim().length()>0){
									dealCommand = dealCommand + strTemp4 + " "; } } } } }
			else if(strFunctionality.trim().toLowerCase().contains("dealreverse")) {
                    String strTemp0 = fnGetExcelData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "DPM_Deal_ID");
                    dealCommand = "/reverse " + strTemp0; }

			System.out.println("[INFO]	Deal Command : "+dealCommand);
			if(dealCommand.trim().length() > 0) {
				// close the command bar if already opened
				if(fnObjectExists(objCmdBarCloseButton, 2) != 0)
					fnObjectClick(objCmdBarCloseButton);
				fnAppSync(2);
				
				if(fnObjectExists(objcmdBar, 2) != 0) {
					fnAppSync(1);
					wDriver.findElement(By.xpath(objcmdBar)).click();
					fnAppSync(1);
					wDriver.findElement(By.xpath(objcmdBar)).clear();
					fnAppSync(1);
					wDriver.findElement(By.xpath(objcmdBar)).sendKeys(dealCommand);
					System.out.println("[INFO]	Successfully Entered Deal Command In Command Bar");
					pageObjects.strGVStepName = "Successfully Entered Deal Command In Command Bar";
					fnAppSync(2);

					if(fnObjectExists(objCmdBarSuggErr, 2) != 0) {
						pageObjects.strScnStatus = false;
						strTxt = wDriver.findElement(By.xpath(objCmdBarSuggErr)).getText();
						System.out.println("[ERROR]	Deal Command Contains Error - "+strTxt);
						pageObjects.strGVErrMsg = "Deal Command Contains Errors  - "+strTxt;
						fnScreenCapture(wDriver, strScreenPath.concat("DealCommandError") + "-"+ dataPointer);
						fnAppSync(1);	

						if(fnObjectExists(objCmdBarCloseButton, 2) != 0)
							fnObjectClick(objCmdBarCloseButton);
						fnAppSync(2);	}
					else {
						fnScreenCapture(wDriver, strScreenPath.concat("DealCommand") + "-"+ dataPointer);
						wDriver.findElement(By.xpath(objcmdBar)).sendKeys(Keys.ENTER);
						fnAppSync(1); 

						for(int intTemp=0;intTemp<10;intTemp++){
							if (fnObjectExists(objIDErrText, 1) != 0) {
								Point objCoordinates = wDriver.findElement(By.xpath(objIDErrText)).getLocation();
								robot.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+87);
								strTxt = wDriver.findElement(By.xpath(objServerErrorContent)).getText();
								robot.mouseMove(objCoordinates.getX()+64, objCoordinates.getY()+83);
								strFound = 0; pageObjects.strScnStatus = false; 
								
								pageObjects.strGVErrMsg = "Failed to Open the " + product + " deal, error message : " + strTxt;
								pageObjects.strGVStepName = "Failed to Book the " + product + " deal, error message : " + strTxt;
								fnScreenCapture(wDriver, strScreenPath.concat("dpmDealBookFailure") + "-"+ dataPointer); }
							else if(fnObjectExists(objBtnBook, 1) != 0) {
								break; } } } }

				if (pageObjects.strScnStatus == true) {
					if (strFunctionality.trim().toLowerCase().contains("dealreverse")) {
						if(fnObjectExists(objDealContinueReverseBtn, 40) != 0) {
							int strIntTemp;
							if(fnObjectExists(objReverseFormReversalReason, 4) != 0) {
								String strTemp = fnGetExcelData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Reversal_Reason");
								if (strTemp.isEmpty() == false){
										strIntTemp = fnSelectValue(objReverseFormReversalReason, strTemp);
										System.out.println("[INFO]   Successfully selected the reversal reason from REVERSE FORM"); } }

							System.out.println("[INFO]   Successfully opened the REVERSE quick deal window");
							pageObjects.strGVStepName = "Successfully opened the REVERSE quick deal window";
							fnScreenCapture(wDriver, strScreenPath.concat("ReverseQckDlWindow") + "-"+ dataPointer); }
						else {
							pageObjects.strScnStatus = false;
							System.out.println("[INFO]	Failed to open the REVERSE quick deal window");
							pageObjects.strGVStepName = "Failed to open the REVERSE quick deal window";
							fnScreenCapture(wDriver, strScreenPath.concat("FailReverseQckDlWindow") + "-"+ dataPointer); } }
					else if (strFunctionality.trim().toLowerCase().contains("dealunwind")) {
						if(fnObjectExists(objContinueUnwindBtn, 40) != 0) {
							System.out.println("[INFO]   Successfully opened the Unwind quick deal window");
							pageObjects.strGVStepName = "Successfully opened the Unwind quick deal window";
							fnScreenCapture(wDriver, strScreenPath.concat("UnWindQckDlWindow") + "-"+ dataPointer); }
						else {
							pageObjects.strScnStatus = false;
							System.out.println("[INFO]	Failed to open the Unwind quick deal window");
							pageObjects.strGVStepName = "Failed to open the Unwind quick deal window";
							fnScreenCapture(wDriver, strScreenPath.concat("FailUnWindQckDlWindow") + "-"+ dataPointer); } }
					else if (strFunctionality.trim().toLowerCase().contains("dealbook")) {
						if(fnObjectExists(objBtnBook, 40) != 0) {
							System.out.println("[INFO]	Successfully opened the quick deal book window");
							pageObjects.strGVStepName = "Successfully opened the quick deal book window";
							fnScreenCapture(wDriver, strScreenPath.concat("QuickDealWindow") + "-"+ dataPointer);

							if (product.trim().toLowerCase().contentEquals("eqo"))
								if (strFunctionality.trim().toLowerCase().contains("dealbook"))
									fnUpdateCmdBarValuesInExcel(dataSheet, dataPointer); }
						else {
							pageObjects.strScnStatus = false;
							System.out.println("[INFO]	Failed to open the quick deal book window");
							pageObjects.strGVStepName = "Failed to open the quick deal book window";
							fnScreenCapture(wDriver, strScreenPath.concat("FailQuickDealWindow") + "-"+ dataPointer); }	} }
				else if (pageObjects.strScnStatus == false) {
					pageObjects.strScnStatus = false;
					System.out.println("[INFO]	Failed to open the quick deal window thru command bar");
					pageObjects.strGVStepName = "Failed to open the quick deal window thru command bar";
					fnScreenCapture(wDriver, strScreenPath.concat("FailQuickDealOpen") + "-"+ dataPointer); } }
			else {
				pageObjects.strScnStatus = false;
				System.out.println("[INFO]	Failed to create the Command bar details....");
				pageObjects.strGVStepName = "Failed to create the Command bar details....";
				fnScreenCapture(wDriver, strScreenPath.concat("QuickCmdBar") + "-"+ dataPointer); } } }
	
// function to confirm the new deals and udpate the excel sheet thru the command bar =============================================================================================================
	public void fnBookDealThroughCmdBar(String dataSheet, String dataPointer, String product) throws Throwable {
		Robot robot = new Robot();
		String strTxt="";
		int strFound = 0;
		if(fnObjectExists(objBtnBook, 40) != 0) {
			fnObjectClick(objBtnBook);
			fnAppSync(1);
			pageObjects.strGVStepName = "Successfully Clicked on Book Button";
			fnScreenCapture(wDriver, strScreenPath.concat("ClickBookButton") + "-"+ dataPointer);

			if(fnObjectExists(objConfirmAction, 2) != 0) {
				fnObjectClick(objConfirmAction);
				pageObjects.strGVStepName = "Successfully Clicked on Confirm Button";
				fnScreenCapture(wDriver, strScreenPath.concat("ConfirmButton") + "-"+ dataPointer);	}

			strFound = 0;
			if(pageObjects.strScnStatus == true) {
				for(int intTemp = 0; intTemp<20; intTemp++) {
					strFound = -1;
					if ((fnObjectExists(objIDSuccessText, 1) != 0) || (fnObjectExists(objIDInBookingText, 1) != 0))
					{ 	String strTxt1="";
						if (fnObjectExists(objIDSuccessText, 1) != 0){
							Point objCoordinates = wDriver.findElement(By.xpath(objIDSuccessText)).getLocation();
							robot.mouseMove(objCoordinates.getX()+63, objCoordinates.getY()+85);
							robot.mouseMove(objCoordinates.getX()+67, objCoordinates.getY()+81);
							strTxt1 = wDriver.findElement(By.xpath(objIDContent)).getText();}
						else if (fnObjectExists(objIDInBookingText, 1) != 0){
							Point objCoordinates = wDriver.findElement(By.xpath(objIDInBookingText)).getLocation();
							robot.mouseMove(objCoordinates.getX()+63, objCoordinates.getY()+85);
							robot.mouseMove(objCoordinates.getX()+67, objCoordinates.getY()+81);
							strTxt1 = wDriver.findElement(By.xpath(objIDInBookingContent)).getText();}
	
						if (strTxt1.trim().toLowerCase().contains("object") == true) {
							strFound = 0;
							strTxt = "[object object]"; }
						else 
							strFound = 1; }

					else if (fnObjectExists(objIDErrText, 1) != 0)
					{	Point objCoordinates = wDriver.findElement(By.xpath(objIDErrText)).getLocation();
						robot.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+87);
						strTxt = wDriver.findElement(By.xpath(objServerErrorContent)).getText();
						robot.mouseMove(objCoordinates.getX()+64, objCoordinates.getY()+83);
						strFound = 0; pageObjects.strScnStatus = false; }
						else if (fnObjectExists(objFormErr, 1) != 0)
						{	strFound = 0; pageObjects.strScnStatus = false;
						strTxt = fnGetObjectAttributeValue(objFormErr, "outerText");}
						if (strFound != -1) break;
						fnAppSync(1); } }

			if ((strFound == 1) && (pageObjects.strScnStatus == true)) {
				boolean strSTxt = false;
				if (fnObjectExists(objIDContent, 1)!=0)
					strTxt = wDriver.findElement(By.xpath(objIDContent)).getText();
				else if (fnObjectExists(objIDInBookingContent, 1)!=0)
					strTxt = wDriver.findElement(By.xpath(objIDInBookingContent)).getText();
				if (strTxt.toLowerCase().contains("object object") == false) {
					if (fnObjectExists(objSimpleNotification, 1) != 0)
						if (fnObjectExists(objIDSuccessText, 1) != 0)
							strSTxt = (wDriver.findElement(By.xpath(objIDSuccessText)).getText().toLowerCase().contains("success"));
						else if (fnObjectExists(objIDInBookingText, 1) != 0)
							strSTxt = (wDriver.findElement(By.xpath(objIDInBookingText)).getText().toLowerCase().contains("in booking"));
					System.out.println("[INFO]   Deal Popup Message details - " + strTxt);	
					if (strSTxt == true) {
						pageObjects.strDealBokID = strTxt; 
						pageObjects.strGVDealID = strTxt;
						pageObjects.strGVStepName = "New Deal Booked Successfully and Deal Id - " + strTxt;
						System.out.println("[INFO]   New Deal Booked Successfully and Deal Id - " + strTxt);
						fnScreenCapture(wDriver, strScreenPath.concat("dpmDealBookSuccess") + "-"+ dataPointer);
						
						fnObjectClick(objPXVLabel);
						
						Point objCoordinates = wDriver.findElement(By.xpath(objIDSuccessText)).getLocation();
						Robot robot1 = new Robot();
						robot1.mouseMove(objCoordinates.getX()+76, objCoordinates.getY()+86);
						robot1.mouseMove(objCoordinates.getX()+71, objCoordinates.getY()+81);
						wDriver.findElement(By.xpath(objIDSuccessText)).click();

						//step to update the deal id into the excel sheet
						fnUpdateExcelCellData(strDataSheetPath, dataSheet, dataPointer, "DPM_Deal_ID", strTxt);
						String strTime = fnGetCurrentDate("yyyy/mm/dd") + " " + fnGetCurrentTime();
						fnUpdateExcelCellData(strDataSheetPath, dataSheet, dataPointer, "Deal_DateTime", strTime); } }
				else
					strFound = 0; }
			if ((strFound == 0) || (pageObjects.strScnStatus == false)) {
				if (fnObjectExists(objServerError, 1) != 0)
					strTxt = wDriver.findElement(By.xpath(objServerErrorContent)).getText();
				pageObjects.strScnStatus = false;
				if (pageObjects.strDealAmend == false) {
					pageObjects.strGVErrMsg = "Failed to Book the " + pageObjects.strProductName + " deal, error message : " + strTxt;
					pageObjects.strGVStepName = "Failed to Book the " + pageObjects.strProductName + " deal, error message : " + strTxt;
					fnScreenCapture(wDriver, strScreenPath.concat("dpmDealBookFailure") + "-"+ dataPointer);
					System.out.println("[ERROR]   Failed to Book the " + pageObjects.strProductName + " deal, error message : " + strTxt); } } } }

// function to enter the command bar details thru the command bar =============================================================================================================
	public void fnUnWindThruCommandBar(String dataSheet, String dataPointer, String product, String strFunctionality) throws Throwable {
		Robot robot = new Robot();
		String strTxt="";
		boolean strSTxt;
		int strFound = 0;
		
		if(fnObjectExists(objContinueUnwindBtn, 40) != 0) {
			//Step to validate the values in UnWind details screen
			pageObjects.strCompareVal = true;
			//step to retrieve the UNWIND FEE from excel and book detail screen and compare
			fnStringObjCompareValues(strDataSheetPath, dataSheet, dataPointer, "Unwind_Fee", "Unwind_Fee", objEqoUnwindFee, "value");
			
			//step to retrieve the UNWIND FEE CURRENCY from excel and book detail screen and compare
			fnStringObjCompareValues(strDataSheetPath, dataSheet, dataPointer, "Unwind_FeeCurr", "Unwind_FeeCurr", objEqoUnwindFeeCurr, "value");
			
			//step to retrieve the FEE SETTLEMENT DATE from excel and book detail screen and compare
			fnDateObjCompareValues(strDataSheetPath, dataSheet, dataPointer, "Unwind_FeeSettlDate", "Unwind_FeeSettlDate", objEqoUnwindFeeSettlDate, "value", "yyyy-mm-dd");
			
			//step to retrieve the FEE BOOK from excel and book detail screen and compare
			fnListObjCompareValues(strDataSheetPath, dataSheet, dataPointer, "Unwind_FeeBook", "Unwind_FeeBook", objEqoUnwindFeeBook, "value");
			
			if (pageObjects.strCompareVal == true) {
				fnObjectClick(objContinueUnwindBtn);
				pageObjects.strGVStepName = "Clicked on the CONTINUE UNWIND Button";
				System.out.println("[INFO]   Clicked on the CONTINUE UNWIND Button");
				fnScreenCapture(wDriver, strScreenPath.concat("dpmDealUnWindSuccess") + "-"+ dataPointer);
				fnAppSync(3);
				if (fnObjectExists(objDealConfirmUnwindBtn, 2)==0) {
					pageObjects.strScnStatus = false;
					System.out.println("[INFO]	Failed to find/ click on the UNWIND button");
					pageObjects.strGVStepName = "Failed to find/ click on the UNWIND button";
					fnScreenCapture(wDriver, strScreenPath.concat("FlUnWindQckDlWindow") + "-"+ dataPointer); }
				else if ((fnObjectExists(objDealConfirmUnwindBtn, 2)!=0)) {
					fnObjectClick(objDealConfirmUnwindBtn);
					fnAppSync(1);				
					String strTxt2 = null;

					if (pageObjects.strScnStatus == true) {
						for(int intTemp = 0; intTemp<180; intTemp++) {
							strFound = -1;
							if (fnObjectExists(objIDPopUpMsg1, 1)!=0){
								Point objCoordinates = wDriver.findElement(By.xpath(objIDPopUpMsg1)).getLocation();
								robot.mouseMove(objCoordinates.getX()+60, objCoordinates.getY()+80);
								robot.mouseMove(objCoordinates.getX()+63, objCoordinates.getY()+82);
								robot.delay(2000);

								if (fnObjectExists(objIDInBookingText, 1) != 0)
								{		String strTxt1="";
										objCoordinates = wDriver.findElement(By.xpath(objIDInBookingText)).getLocation();
										robot.mouseMove(objCoordinates.getX()+60, objCoordinates.getY()+80);
										robot.mouseMove(objCoordinates.getX()+63, objCoordinates.getY()+82);
										robot.delay(2000);
										strTxt = wDriver.findElement(By.xpath(objIDInBookingContent)).getText();
									
									if (strTxt.trim().toLowerCase().contains("object") == true) {
										strFound = 0;
										strTxt = "[object object]"; }
									else {
										if (strTxt.trim().toLowerCase().contains(":")){
											String[] strTemp = strTxt.split(":");
											strTxt = strTemp[1].trim();
											strFound = 1;
											strSTxt = true; } } }
								else if (fnObjectExists(objIDSuccessText, 1) != 0)
								{ 		String strTxt1="";
										objCoordinates = wDriver.findElement(By.xpath(objIDSuccessText)).getLocation();
										robot.mouseMove(objCoordinates.getX()+63, objCoordinates.getY()+85);
										robot.mouseMove(objCoordinates.getX()+67, objCoordinates.getY()+81);
										strTxt = wDriver.findElement(By.xpath(objIDContent)).getText();
																	
									if (strTxt.trim().toLowerCase().contains("object") == true) {
										strFound = 0;
										strTxt = "[object object]"; }
									else {
										if (strTxt.trim().toLowerCase().contains(":")){
											String[] strTemp = strTxt.split(":");
											strTxt = strTemp[1].trim();
											strFound = 1;
											strSTxt = true; } 
										else {
											strFound = 1;
											strSTxt = true;	} } }
			
								else if (fnObjectExists(objIDErrText, 1) != 0)
								{	objCoordinates = wDriver.findElement(By.xpath(objIDErrText)).getLocation();
									robot.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+87);
									strTxt = wDriver.findElement(By.xpath(objServerErrorContent)).getText();
									robot.mouseMove(objCoordinates.getX()+64, objCoordinates.getY()+83);
									strFound = 0; pageObjects.strScnStatus = false; 
									strTxt = fnGetObjectAttributeValue(objServerErrorContent, "outerText"); }
								else if (fnObjectExists(objFormErr, 1) != 0)
								{	strFound = 0; pageObjects.strScnStatus = false;
									strTxt = fnGetObjectAttributeValue(objFormErr, "outerText"); }
								else {
									String strNewID = fnGetObjectAttributeValue(objBkDealHeaderDetails, "outerText");
									if (strNewID.trim().toLowerCase().contentEquals(strBkDealID.trim().toLowerCase())==false){
										if (strNewID.trim().toLowerCase().contains("in booking")) {
											String[] strList = strNewID.split(" ");
											if (strList.length>1) {
												strTxt = strList[0].trim(); 
												strFound = 1;
												System.out.println("[INFO]    BASIS Deal ID : " + strTxt);} } } }	 }	
							if (strFound != -1) break; } } }
				
				if ((strFound == 1) && (pageObjects.strScnStatus == true)) {
					strSTxt = false;

					if (fnObjectExists(objIDContent, 1)!=0)
						strTxt = wDriver.findElement(By.xpath(objIDContent)).getText();
					else if (fnObjectExists(objIDInBookingContent, 1)!=0)
						strTxt = wDriver.findElement(By.xpath(objIDInBookingContent)).getText();
					
					if (strTxt.trim().toLowerCase().contains(":")){
						String[] strTemp = strTxt.split(":");
						strTxt = strTemp[1].trim();	}

					if (strTxt.toLowerCase().contains("object object") == false) {
						if (fnObjectExists(objSimpleNotification, 1) != 0)
							if (fnObjectExists(objIDSuccessText, 1)!=0)
								strSTxt = (wDriver.findElement(By.xpath(objIDSuccessText)).getText().toLowerCase().contains("success"));
							else if (fnObjectExists(objIDInBookingText, 1)!=0)
								strSTxt = (wDriver.findElement(By.xpath(objIDInBookingText)).getText().toLowerCase().contains("in booking"));
						
							System.out.println("[INFO]   Deal Popup Message details - " + strTxt);	
							if (strSTxt == true) {
								pageObjects.strGVStepName = "Deal UnWind Successfull and Deal Id - " + strTxt;
								System.out.println("[INFO]   Deal UnWind Successfull and Deal Id - " + strTxt);
								pageObjects.strDealBokID = strTxt;
								pageObjects.strGVDealID = strTxt;
								fnScreenCapture(wDriver, strScreenPath.concat("dpmDealBookSuccess") + "-"+ dataPointer);
								
								fnObjectClick(objPXVLabel);
								if (fnObjectExists(objIDSuccessText, 1)!=0){
									Point objCoordinates = wDriver.findElement(By.xpath(objIDSuccessText)).getLocation();
									Robot robot1 = new Robot();
									robot1.mouseMove(objCoordinates.getX()+76, objCoordinates.getY()+86);
									robot1.mouseMove(objCoordinates.getX()+71, objCoordinates.getY()+81);
									wDriver.findElement(By.xpath(objIDSuccessText)).click();}
								
								//step to update the deal id into the excel sheet
								fnUpdateExcelCellData(strDataSheetPath, dataSheet, dataPointer, "DPM_Deal_ID", strTxt);
								String strTime = fnGetCurrentDate("yyyy/mm/dd") + " " + fnGetCurrentTime();
								fnUpdateExcelCellData(strDataSheetPath, dataSheet, dataPointer, "Deal_DateTime", strTime); } }
					else
						strFound = 0; }
				else if ((strFound == -1) || (strFound == 0) || (pageObjects.strScnStatus == false)) {
					pageObjects.strScnStatus = false;
					pageObjects.strGVErrMsg = "Failed to Unwind the " + pageObjects.strProductName + " deal, error msg - " + strTxt;
					pageObjects.strGVStepName = "Failed to Unwind the " + pageObjects.strProductName + " deal, error msg - " + strTxt;
					fnScreenCapture(wDriver, strScreenPath.concat("dpmDealBookFailure") + "-"+ dataPointer);
					System.out.println("[ERROR]   Failed to Unwind the " + pageObjects.strProductName + " deal, error msg - " + strTxt); } }
			else {
				pageObjects.strScnStatus = false;
				System.out.println("[INFO]	Failed to validate the details in the deal Unwind window");
				pageObjects.strGVStepName = "Failed to validate the details in the deal Unwind window";
				fnScreenCapture(wDriver, strScreenPath.concat("FlUnWindQckDlWindow") + "-"+ dataPointer); } }
		else {
			pageObjects.strScnStatus = false;
			System.out.println("[INFO]	Failed to open the Unwind quick deal window");
			pageObjects.strGVStepName = "Failed to open the Unwind quick deal window";
			fnScreenCapture(wDriver, strScreenPath.concat("FailUnWindQckDlWindow") + "-"+ dataPointer);	} }

// function to reverse deal through command bar===============================================================================================
        public void fnReverseDealFromCmdBar(String strDataSheet, String strDataPointer) throws Throwable{

              //step to check for the CONTINUE REVERSE button and proceed
              if (fnObjectExists(objDealContinueReverseBtn, 10) != 0) {

                    System.out.println("[INFO]   Clicked on Continue Reverse button for Deal ID - "+pageObjects.strGVDealID);
                    pageObjects.strGVStepName = "Clicked on Continue Reverse for Deal ID - "+pageObjects.strGVDealID;
                    fnScreenCapture(wDriver, strScreenPath.concat("dpmDealReverseContinueBtn") + "-" + strDataPointer);
                    fnObjectClick(objDealContinueReverseBtn);

                    //step to check for the CONFIRM button and proceed
                    if (fnObjectExists(objDealConfirmReverseBtn, 10) != 0) {

                          System.out.println("[INFO]   Clicked on Confirm button for Deal ID - "+pageObjects.strGVDealID);
                          pageObjects.strGVStepName = "Clicked on Confirm for Deal ID - "+pageObjects.strGVDealID;
                          fnScreenCapture(wDriver, strScreenPath.concat("dpmDealReverseConfirmBtn") + "-" + strDataPointer);

                          fnObjectClick(objDealConfirmReverseBtn);
                          int strFound = -1;
                          Robot robot = new Robot();
                          String strTxt;
                          boolean strSTxt = false;
                          for(int intTemp = 0; intTemp<intMsgWaitTime; intTemp++) {
          					if (fnObjectExists(objIDPopUpMsg, 1)!=0){
          						Point objCoordinates = wDriver.findElement(By.xpath(objIDPopUpMsg)).getLocation();
          						robot.mouseMove(objCoordinates.getX()+60, objCoordinates.getY()+80);
          						robot.mouseMove(objCoordinates.getX()+63, objCoordinates.getY()+82);
          						robot.delay(2000);

          						if (fnObjectExists(objIDInBookingText, 1) != 0)
          						{		String strTxt1="";
          								objCoordinates = wDriver.findElement(By.xpath(objIDInBookingText)).getLocation();
          								robot.mouseMove(objCoordinates.getX()+60, objCoordinates.getY()+80);
          								robot.mouseMove(objCoordinates.getX()+63, objCoordinates.getY()+82);
          								robot.delay(2000);
          								strTxt = wDriver.findElement(By.xpath(objIDInBookingContent)).getText();
          							
          							if (strTxt.trim().toLowerCase().contains("object") == true) {
          								strFound = 0;
          								strTxt = "[object object]"; }
          							else {
          								if (strTxt.trim().toLowerCase().contains(":")){
          									String[] strTemp = strTxt.split(":");
          									strTxt = strTemp[1].trim();
          									strFound = 1;
          									strSTxt = true; } } }
          						else if (fnObjectExists(objIDSuccessText, 1) != 0)
          						{ 		String strTxt1="";
          								objCoordinates = wDriver.findElement(By.xpath(objIDSuccessText)).getLocation();
          								robot.mouseMove(objCoordinates.getX()+63, objCoordinates.getY()+85);
          								robot.mouseMove(objCoordinates.getX()+67, objCoordinates.getY()+81);
          								strTxt = wDriver.findElement(By.xpath(objIDContent)).getText();
          															
          							if (strTxt.trim().toLowerCase().contains("object") == true) {
          								strFound = 0;
          								strTxt = "[object object]"; }
          							else {
          								if (strTxt.trim().toLowerCase().contains(":")){
          									String[] strTemp = strTxt.split(":");
          									strTxt = strTemp[1].trim();
          									strFound = 1;
          									strSTxt = true; } 
          								else {
          									strFound = 1;
          									strSTxt = true;	} } }
          	
          						else if (fnObjectExists(objIDErrText, 1) != 0)
          						{	objCoordinates = wDriver.findElement(By.xpath(objIDErrText)).getLocation();
          							robot.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+87);
          							strTxt = wDriver.findElement(By.xpath(objServerErrorContent)).getText();
          							robot.mouseMove(objCoordinates.getX()+64, objCoordinates.getY()+83);
          							strFound = 0; pageObjects.strScnStatus = false; 
          							strTxt = fnGetObjectAttributeValue(objServerErrorContent, "outerText"); }
          						else if (fnObjectExists(objFormErr, 1) != 0)
          						{	strFound = 0; pageObjects.strScnStatus = false;
          							strTxt = fnGetObjectAttributeValue(objFormErr, "outerText"); }
          						else {
          							String strNewID = fnGetObjectAttributeValue(objBkDealHeaderDetails, "outerText");
          							if (strNewID.trim().toLowerCase().contentEquals(strBkDealID.trim().toLowerCase())==false){
          								if (strNewID.trim().toLowerCase().contains("in booking")) {
          									String[] strList = strNewID.split(" ");
          									if (strList.length>1) {
          										strTxt = strList[0].trim(); 
          										strFound = 1;
          										System.out.println("[INFO]    BASIS Deal ID : " + strTxt);} } } }	 }	
          					if (strFound != -1) break; }

                          if (strFound == 1) {
                                try {
                                	if (fnObjectExists(objIDContent, 1)!=0)
                                		pageObjects.strGVReverseDealID = fnGetObjectAttributeValue(objIDContent, "innerText");
                                	else if (fnObjectExists(objIDInBookingContent, 1)!=0)
                                		pageObjects.strGVReverseDealID = fnGetObjectAttributeValue(objIDInBookingContent, "innerText"); }
                                catch (Exception exp) {
                                    fnAppSync(3);
                                    if (fnObjectExists(objIDContent, 1)!=0)
                                  		pageObjects.strGVReverseDealID = fnGetObjectAttributeValue(objIDContent, "innerText");
                                  	else if (fnObjectExists(objIDInBookingContent, 1)!=0)
                                  		pageObjects.strGVReverseDealID = fnGetObjectAttributeValue(objIDInBookingContent, "innerText"); }
                                
                                if (pageObjects.strGVReverseDealID.trim().toLowerCase().contains(":")){
                                	String[] strTemp = pageObjects.strGVReverseDealID.split(":");
                                	pageObjects.strGVReverseDealID = strTemp[1].trim(); }
                                
                                pageObjects.strGVStepName = "Successfully executed the Reveser & NEW Deal Reverse ID - " + pageObjects.strGVReverseDealID;
                                fnScreenCapture(wDriver, strScreenPath.concat("dpmDealReverseSuccess") + "-" + strDataPointer);
                                if (pageObjects.strScnStatus == true){
                                      if (fnObjectExists(objIDSuccessText, 1)!=1) {
                                            Point objCoordinates = wDriver.findElement(By.xpath(objIDSuccessText)).getLocation();
                                            Robot robot1 = new Robot();
                                            robot1.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+86);
                                            fnAppSync(1);
                                            robot1.mouseMove(objCoordinates.getX()+69, objCoordinates.getY()+80);
                                            fnObjectClick(objIDSuccessText); }
                                      else if (fnObjectExists(objIDInBookingText, 1)!=1) {
                                          Point objCoordinates = wDriver.findElement(By.xpath(objIDInBookingText)).getLocation();
                                          Robot robot1 = new Robot();
                                          robot1.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+86);
                                          fnAppSync(1);
                                          robot1.mouseMove(objCoordinates.getX()+69, objCoordinates.getY()+80);
                                          fnObjectClick(objIDInBookingText); } 

                                      System.out.println("[INFO]   Successfully Completed Reverse process for the newly booked deal ID - " + pageObjects.strGVReverseDealID);
                                      fnUpdateExcelCellData(pageObjects.strDataSheetPath, strDataSheet, strDataPointer, "DPM_Deal_ID", pageObjects.strGVReverseDealID); }
                                else {
                                      pageObjects.strScnStatus = false;
                                      System.out.println("[ERROR]   Failed to find the CONTINUE REVERSE button in the deal book screen");
                                      pageObjects.strGVStepName = "Failed to find the CONTINUE REVERSE button in the deal book screen";
                                      fnScreenCapture(wDriver, strScreenPath.concat("dpmDealContinueReverseDealFailure") + "-" + strDataPointer);
                                      pageObjects.strGVErrMsg = "Failed to find the CONTINUE REVERSE button in the deal book screen"; } 
                          }else {
                                pageObjects.strScnStatus = false;
                                System.out.println("[ERROR]   Failed to REVERSE the deal - ID : " + pageObjects.strGVDealID);
                                pageObjects.strGVStepName = "Failed to REVERSE the deal - ID : " + pageObjects.strGVDealID;
                                fnScreenCapture(wDriver, strScreenPath.concat("dpmDealContinueReverseDealFailure") + "-" + strDataPointer);
                                pageObjects.strGVErrMsg = "Failed to REVERSE the deal - ID : " + pageObjects.strGVDealID; } 
                    }else {
                          pageObjects.strScnStatus = false;
                          System.out.println("[ERROR]   Failed to find the REVERSE CONFIRM button in the deal book screen");
                          pageObjects.strGVStepName = "Failed to find the REVERSE CONFIRM button in the deal book screen";
                          fnScreenCapture(wDriver, strScreenPath.concat("dpmDealContinueReverseDealFailure") + "-" + strDataPointer);
                          pageObjects.strGVErrMsg = "Failed to find the REVERSE CONFIRM button in the deal book screen";
                          if (fnObjectExists(objDealCancelBtn, 3) != 0)
                                try {
                                      fnObjectClick(objDealCancelBtn); }
                          catch (Exception expErr) {
                                System.out.println("[ERROR]   Unable to click on the Cancel Button"); } }
              }else {
                    pageObjects.strScnStatus = false;
                    System.out.println("[ERROR]   Failed to find the CONTINUE REVERSE button in the deal book screen");
                    pageObjects.strGVStepName = "Failed to find the CONTINUE REVERSE button in the deal book screen";
                    fnScreenCapture(wDriver, strScreenPath.concat("dpmDealContinueReverseDealFailure") + "-" + strDataPointer);
                    pageObjects.strGVErrMsg = "Failed to find the CONTINUE REVERSE button in the deal book screen"; } }	

//Function to retrive the values from the DB ==========================================================================================
	public static String fnGetDBRowValues(String strDBQuery, String strEnv){
		String uName = "mtm_oper";
		String pwd = "mtm_oper1";
		String url = "";
		String strValues = "";

		if ((strEnv.trim().toLowerCase().contentEquals("uat")) || (strEnv.trim().toLowerCase().contentEquals("uat1")))
			url = "jdbc:sybase:Tds:172.20.184.78:7900";
		else if ((strEnv.trim().toLowerCase().contains("azure")) || (strEnv.trim().toLowerCase().contentEquals("ist4")))
			url = "jdbc:sybase:Tds:172.22.141.115:4010";
		
		System.out.println(strDBQuery);
		try { Class.forName("com.sybase.jdbc.SybDriver"); }
	    catch (ClassNotFoundException cnfe) { }
		
		try { System.out.println("Opening a connection.");			 
			Connection con = DriverManager.getConnection(url, uName, pwd);
			Statement querySet = con.createStatement();
			System.out.println("Successfully opened connection.");
			ResultSet rs = querySet.executeQuery(strDBQuery);
			ResultSetMetaData rsmd = rs.getMetaData();
			System.out.println("Successfully executed the query.");
			int intTemp = 0;	

			while(rs.next()) {
				intTemp = intTemp + 1;
				strValues = rs.getString(1).trim() + ";" + rs.getString(2).trim();
				System.out.println("DB Table Value : " + strValues);	}

			rs.close();
			querySet.close();
			con.close(); }
		catch (SQLException sqe) {
			System.out.println("Failed to get the DB connection.");
			System.out.println("Unexpected exception : " +
				   sqe.toString() + ", sqlstate = " +
				   sqe.getSQLState()); 
			strValues = "DB-ERROR";}
		return strValues; }	

//Function to create a file with the text with the given content =============================================================================================================
	public static void fnWriteTxtFile(String strFileName, String strFileText) throws Throwable{
		String outputFileName = strFileName.trim();
		System.out.println("[INFO]   File to create : " + outputFileName);
		FileWriter fileWriter = null;
		try{
			fileWriter = new FileWriter(outputFileName);
			fileWriter.append(strFileText);
			fileWriter.flush();
			fileWriter.close(); }
		catch (Exception exp){
			System.out.println("[ERROR]    Failed to write to file -- " + strFileName);	} }	

//Function verify the given status in the exported deals =====================================================================================================================
	public int fnVerifyExportDeals(String strFileName, String strVal) throws Throwable {
		FileInputStream fstream = new FileInputStream(strFileName);
		@SuppressWarnings("resource")
		BufferedReader br = new BufferedReader(new InputStreamReader(fstream));
		String strLine;
		String[] strListVal = strVal.split(";");
		int intFound = 0, intLineCnt = 1;
		pageObjects.strBulkDealId = "";
		//Read File Line By Line
		while ((strLine = br.readLine()) != null) {
			if (intLineCnt > 1) {
				for(int intTemp=0;intTemp<strListVal.length;intTemp++) {
					if (strLine.trim().toLowerCase().contains(strListVal[intTemp].trim().toLowerCase())==false)
						intFound = 1; }	}
			intLineCnt++;
			if (intFound == 1) {
				break; }
			else {
				String strLine1 = strLine.replace("\"", "");
				String[] strListString = strLine1.split(",");
				pageObjects.strBulkDealId = pageObjects.strBulkDealId + ";" + strListString[0]; } }
		System.out.println(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>..." + pageObjects.strBulkDealId);
		strBulkDealId = strBulkDealId.substring(1, pageObjects.strBulkDealId.length());
		pageObjects.strBulkDealId = pageObjects.strBulkDealId.substring(pageObjects.strBulkDealId.indexOf(";") + 1, pageObjects.strBulkDealId.length());
		return intFound; }

//Function to retrive the values from the DB ==========================================================================================
	public static String fnGetDataFromDB(String strDBQuery, String strEnv) {
		String uName = "mtm_oper";
		String pwd = "mtm_oper1";
		String url = "";
		String strValues = "";

		if ((strEnv.trim().toLowerCase().contentEquals("uat")) || (strEnv.trim().toLowerCase().contentEquals("uat1")))
			url = "jdbc:sybase:Tds:172.20.184.78:7900";
		else if ((strEnv.trim().toLowerCase().contains("azure")) || (strEnv.trim().toLowerCase().contentEquals("ist4")))
			url = "jdbc:sybase:Tds:172.22.141.115:4010";
		
		System.out.println(strDBQuery);
		try { Class.forName("com.sybase.jdbc.SybDriver"); }
	    catch (ClassNotFoundException cnfe) { }
		
		try { System.out.println("Opening a connection.");			 
			Connection con = DriverManager.getConnection(url, uName, pwd);
			Statement querySet = con.createStatement();
			System.out.println("Successfully opened connection.");
			ResultSet rs = querySet.executeQuery(strDBQuery);
			ResultSetMetaData rsmd = rs.getMetaData();
			System.out.println("Successfully executed the query.");
			int intTemp = 0;	

			while(rs.next()) {
				intTemp = intTemp + 1;
				strValues = strValues + ";" + rs.getString(1).trim();
				System.out.println("DB Table Value : " + strValues);	}
			
			strValues = strValues.substring(1, strValues.length());
			rs.close();
			querySet.close();
			con.close(); }
		catch (SQLException sqe) {
			System.out.println("Failed to get the DB connection.");
			System.out.println("Unexpected exception : " +
				   sqe.toString() + ", sqlstate = " +
				   sqe.getSQLState()); 
			strValues = "DB-ERROR";}
		return strValues; }
	
	//Function to highlight the object on the browser with a boarder line/ color =============================================================================================
	public static void fnHighlightMe(WebDriver driver,WebElement element) throws InterruptedException{
		   JavascriptExecutor js = (JavascriptExecutor)driver;
		   for (int iCnt = 0; iCnt < 3; iCnt++) {
		         js.executeScript("arguments[0].style.border='4px groove green'", element);
		         Thread.sleep(1000);
		         js.executeScript("arguments[0].style.border=''", element); } }

//Function to add/ subtract the no. of days from the given date in yyyy-mm-dd format ===============================================================================
	public String fnDateSumDuff(String strDate, String dateOp, int intDays) throws ParseException {
		String dt = "2017-03-15";  // Given date
    	SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
    	Calendar c = Calendar.getInstance();
    	c.setTime(sdf.parse(dt));
    	if (dateOp.trim().toLowerCase().contentEquals("add"))
    		c.add(Calendar.DATE, intDays);  // number of days to add
    	else if (dateOp.trim().toLowerCase().contentEquals("sub"))
    		c.add(Calendar.DATE, intDays);  // number of days to add
    	dt = sdf.format(c.getTime());    	
    	System.out.println(dt);
		return strDate;	 }

// function to format date =======================================================================================
	public String fnConvertDateFormat(String date, String month, String year, String outputFormat) throws Throwable{
	           try{
	                // format and replace the date
	        	   outputFormat = outputFormat.trim().toLowerCase();
	                if(date.trim().length() == 1){
	                     date = "0"+date;
	                     outputFormat = outputFormat.replace("d", date); }
	                outputFormat = outputFormat.replace("dd", date);

	                // format and replace the month
	                int monthCount = StringUtils.countMatches(outputFormat.trim().toLowerCase(), "m");
	                if(monthCount == 3){
	                     if(month.trim().length() <= 2){
	                           if(month.trim().length() == 1) {
	                                month = "0"+month; }
	                           switch (month) {
	                           case "01": month = "JAN";break;
	                           case "02": month = "FEB";break;
	                           case "03": month = "MAR";break;
	                           case "04": month = "APR";break;
	                           case "05": month = "MAY";break;
	                           case "06": month = "JUN";break;
	                           case "07": month = "JUL";break;
	                           case "08": month = "AUG";break;
	                           case "09": month = "SEP";break;
	                           case "10": month = "OCT";break;
	                           case "11": month = "NOV";break;
	                           case "12": month = "DEC";break;
	                           default:break; }  }
	                     outputFormat = outputFormat.replace("mmm", month);
	                }else if(monthCount == 2){
	                     if(month.trim().length() == 1){
	                           month = "0"+month;
	                     }else if(month.trim().length() == 3){
	                           switch (month.toUpperCase()) {
	                           case "JAN": month = "01";break;
	                           case "FEB": month = "02";break;
	                           case "MAR": month = "03";break;
	                           case "APR": month = "04";break;
	                           case "MAY": month = "05";break;
	                           case "JUN": month = "06";break;
	                           case "JUL": month = "07";break;
	                           case "AUG": month = "08";break;
	                           case "SEP": month = "09";break;
	                           case "OCT": month = "10";break;
	                           case "NOV": month = "11";break;
	                           case "DEC": month = "12";break;
	                           default:break; } }
	                     outputFormat = outputFormat.replace("mm", month); }

	                // format and replace the year ==========================================================
	                //String strYear = outputFormat.substring(outputFormat.length()-4, outputFormat.length()).trim();
	                //int yearCount = StringUtils.countMatches(strYear.trim().toLowerCase(), "y");
	                	                
	               if(outputFormat.trim().toLowerCase().contains("yyyy")){
	                     if(year.trim().length() == 2){
	                           year = "20"+year; }
	                     outputFormat = outputFormat.trim().toLowerCase().replace("yyyy", year); }
	               else  if(outputFormat.trim().toLowerCase().contains("yy")){
	                     if(year.trim().length() == 4){
	                           year = year.substring(2, 2); }
	                     outputFormat = outputFormat.replace("yy", year); }
	           }catch(Exception e){
	                System.out.println("Exception Occured While formatting Date format - "+e.getMessage());
	                outputFormat = ""; }
	           //System.out.println("CONVERT DATE : " + outputFormat);
	           return outputFormat; }

//Function to read the txt files data and count the no. of lines in the given text file ============================================================================
	@SuppressWarnings("resource")
	public static int fnTxtFileLinesCount(String strInputFileName) throws Throwable {
		String strLine="";
		int intOneLineCount = -1;
		File fileStream = new File(strInputFileName);
		if(fileStream.exists() && !fileStream.isDirectory()) {
			FileInputStream fstream = new FileInputStream(strInputFileName);
			BufferedReader br = new BufferedReader(new InputStreamReader(fstream));
			intOneLineCount = 0;
			while ((strLine = br.readLine()) != null)
				intOneLineCount = intOneLineCount + 1; }
		return intOneLineCount; }

//Function to compare the period data extracted text files  ==========================================================================================================
	@SuppressWarnings("resource")
	public static int fnPeriodDataCSVFileCompare(String strInputFileName1, String strInputFileName2) throws Throwable {		
		String strLine;
		int intFileCompare = 0;
		File fileStream1 = new File(strInputFileName1);
		File fileStream2 = new File(strInputFileName2);
		if ((fileStream1.exists() && !fileStream1.isDirectory()) && (fileStream2.exists() && !fileStream2.isDirectory())) {
			
			//Step to get the lines count, define the array for file 1 and read the data into the array =============================================
			int intLnCnt = fnTxtFileLinesCount(strInputFileName1);
			String[] strLine1 = new String[intLnCnt];
			intLnCnt = 0;
			FileInputStream fstream1 = new FileInputStream(strInputFileName1);
			BufferedReader br1 = new BufferedReader(new InputStreamReader(fstream1));
			while ((strLine = br1.readLine()) != null) {
				strLine1[intLnCnt] = strLine.trim();
				intLnCnt = intLnCnt + 1; }

			//Step to get the lines count, define the array for file 2 and read the data into the array =============================================			
			intLnCnt = fnTxtFileLinesCount(strInputFileName2);
			String[] strLine2 = new String[intLnCnt];
			intLnCnt = 0;
			FileInputStream fstream2 = new FileInputStream(strInputFileName2);
			BufferedReader br2 = new BufferedReader(new InputStreamReader(fstream2));
			while ((strLine = br2.readLine()) != null) {
				strLine2[intLnCnt] = strLine.trim();
				intLnCnt = intLnCnt + 1; }

			for(int intTemp=0;intTemp<strLine1.length;intTemp++) {
				String[] strString1 = strLine1[intTemp].split(";");
				String[] strString2 = strLine2[intTemp].split(";");
				int intValCompare = 0;
				
				for(int intCol = 0; intCol < strString1.length; intCol++) {
					String strVal1 = strString1[intCol];
					String strVal2 = strString2[intCol];
					
					if (strVal1.trim().contentEquals(strVal2.trim())==false) {
						if ((strVal1.trim().substring(0, 1).contentEquals("-")==true) && (strVal1.trim().substring(0, 1).contentEquals("-")==false)) {
							
						} 
						else if ((strVal1.trim().substring(0, 1).contentEquals("-")==false) && (strVal1.trim().substring(0, 1).contentEquals("-")==true)) {

						} else {
							intValCompare = 1;	} } }

				if (intValCompare == 1) {
					intFileCompare = 1;
					pageObjects.strScnStatus = false;
					System.out.println("[ERROR]   CDS Period Data Comparison Failed, Row : " + (intTemp+1) + ", Deal 1 : " + strLine1[intTemp] + " <<>> Deal 2 : " + strLine2[intTemp]); } } }
		return intFileCompare; }	

	// function to add / sub days ==============================================================================================
	public String fnAddOrSubWorkingDays(String inputDate, int noOfDays, String addOrSub) throws Throwable {
	           String output = "";
	           DateTimeFormatter formatter = null;
	           LocalDate localDate = null;
	           try{
	                formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");
	                localDate = LocalDate.parse(inputDate, formatter);

	                if (noOfDays == 0) {
	                     return localDate.format(formatter);
	                }else{
	                     int addedDays = 0;
	                     while (addedDays < noOfDays) {
	                           if(addOrSub.trim().toLowerCase().contentEquals("add")){
	                                localDate = localDate.plusDays(1);
	                           }else if(addOrSub.trim().toLowerCase().contentEquals("sub")){
	                                localDate = localDate.minusDays(1); }
	                           if (!(localDate.getDayOfWeek() == DayOfWeek.SATURDAY ||
	                                     localDate.getDayOfWeek() == DayOfWeek.SUNDAY)) {
	                                ++addedDays; } } }
	           }catch(Exception e){
	                System.out.println("Exception Occured While adding / substracting the working days - "+e.getMessage());
	                return output; }
	           System.out.println("[INFO]    Output Date : " + localDate.format(formatter));
	           return localDate.format(formatter); }

//function to enter the data into the input object field =========================================================================================
	public void fnSetObjInputData(String dataSheet, String dataPointer, String strObj, String strColName) throws Throwable{
		String strVal="";
		Robot robot = new Robot();
		strVal = fnGetExcelData(strDataSheetPath, dataSheet, dataPointer, strColName).trim();
		if (strVal.isEmpty() == false) {
			if (fnObjectExists(strObj, 10) != 0) {
				try{
					wDriver.findElement(By.xpath(strObj)).click();
					wDriver.findElement(By.xpath(strObj)).clear();
					robot.delay(500);
					wDriver.findElement(By.xpath(strObj)).sendKeys(strVal);
					robot.delay(100);
					robot.keyPress(KeyEvent.VK_ENTER);
					robot.keyRelease(KeyEvent.VK_ENTER);
					robot.delay(100); }
				catch (Exception exp){
					System.out.println("[ERROR]    Unable to input the data in object : " + strObj); } }
		robot.delay(100); }
		else {
			strVal = fnGetObjectAttributeValue(strObj, "value");
			fnUpdateExcelCellData(strDataSheetPath, dataSheet, dataPointer, strColName, strVal); } }

//function to select the data into the drop down object field =========================================================================================
	public void fnSelectObjDropDownData(String dataSheet, String dataPointer, String strObj, String strColName, String strValCaseType) throws Throwable{
		String strVal="";
		int strIntTemp=0;
		Robot robot = new Robot();
		strVal = fnGetExcelData(strDataSheetPath, dataSheet, dataPointer, strColName).trim();
		if (strVal.isEmpty() == false) {
			if (strValCaseType.trim().contentEquals("lower"))
				strVal = strVal.trim().toLowerCase();
			if (fnObjectExists(strObj, 10) != 0)
				strIntTemp = fnSelectValue(strObj, strVal.trim()); }
		else {
			strVal = fnGetObjectAttributeValue(strObj, "value");
			fnUpdateExcelCellData(strDataSheetPath, dataSheet, dataPointer, strColName, strVal); }
		robot.delay(400); }
	

//function to SELECT DROP DOWN VALUE - KEY IN the drop down object field =========================================================================================
	public void fnSelectObjDrpDwnEnterVal(String dataSheet, String dataPointer, String strObj, String strColName, String strValCaseType) throws Throwable{
		String strVal="";
		int strIntTemp=0;
		Robot robot = new Robot();
		strVal = fnGetExcelData(strDataSheetPath, dataSheet, dataPointer, strColName).trim();
		if (strVal.isEmpty() == false) {
			if (strValCaseType.trim().contentEquals("lower"))
				strVal = strVal.trim().toLowerCase();
			if (fnObjectExists(strObj, 10) != 0){
				fnObjectClick(strObj);
				wDriver.findElement(By.xpath(strObj)).sendKeys(strVal);
				robot.delay(100);
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
				
				//strIntTemp = fnSelectValue(strObj, strVal.trim()); 
				} }
		else {
			strVal = fnGetObjectAttributeValue(strObj, "value");
			fnUpdateExcelCellData(strDataSheetPath, dataSheet, dataPointer, strColName, strVal); }
		robot.delay(400); }
	
// Function to read the complete excel sheet and initialize it in 2D array ======================================================
	@SuppressWarnings({ "deprecation", "resource" })
	public static void fnReadExcelSheet(String fileName, String sheetName) throws Throwable {
		String cellValue = "";
		try{
			File file = new File(fileName);
			if(file.exists()){
				InputStream myxls = new FileInputStream(fileName);
				HSSFWorkbook book = new HSSFWorkbook(myxls);
				HSSFSheet sheet = book.getSheet(sheetName);
				int rows = fnGetExcelRowCount(fileName, sheetName);
				int columns = fnGetExcelColumnCount(fileName, sheetName);
				System.out.println("Rows : " + rows + ", Columns : " + columns);
				pageObjects.excelValues = new String[rows][columns];
				for(int i = 0; i < rows ; i++){
					Row row = sheet.getRow(i);
					for(int j = 0 ; j < row.getLastCellNum() ; j++){
						//System.out.println(i+" , "+j);
						row = sheet.getRow(i);
						Cell cell = row.getCell(j);
						if(cell == null)
							cellValue = "";
						else if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
							cellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
						else if(cell.getCellType() == Cell.CELL_TYPE_STRING)
							cellValue = cell.getStringCellValue();
						else if(cell.getCellType() == Cell.CELL_TYPE_BLANK)
							cellValue = "";
						//System.out.println(cellValue);
						pageObjects.excelValues[i][j] = cellValue;	} } } }
		catch(Exception e){
			System.out.println("Exception Occured | comlib | fnReadExcelSheet - "+e.getMessage());}	}

// Function to find the indexes and return the value ============================================================================
	public static String fnGetExcelCellValueUsingIndex(String excelValues[][],String dataPointer, String columnName)throws Throwable {
		String cellValue = "";

		int rowIndex = -1;
		int columnIndex = -1;
		try{
			for(int row = 0 ; row < excelValues.length ; row++){
				if(excelValues[row][0].trim().toLowerCase().contentEquals(dataPointer.trim().toLowerCase())){
					rowIndex = row;
					break;}}
			for(int column = 0 ; column < (excelValues[0].length) ; column++){
				if(excelValues[0][column].trim().toLowerCase().contentEquals(columnName.trim().toLowerCase())){
					columnIndex = column;
					break;}}
			System.out.println("Row Index : "+rowIndex);
			System.out.println("Column Index : "+columnIndex);
			cellValue = excelValues[rowIndex][columnIndex]; }
		catch(Exception e) {
			rowIndex = rowIndex + 0;
			columnIndex = columnIndex + 0;
			cellValue = cellValue + "";		}
		System.out.println("Excel - Array - columnName : " + cellValue);
		return cellValue;	}

//Function to compute the date and update in the excel cell =====================================================================================
	public String fnDateComputeUpdateCell(String strObj,String strDate,String strDataSheet,String strDataPointer,String strColName) throws Throwable{
		String strVal="0";
		if (strDate.trim().toLowerCase().contentEquals("today")) {
			strVal = fnGetObjectAttributeValue(strObj, "value");
			if (strVal.trim().contains("|")){
				String[] strVal1 = strVal.split("\\|");
				strVal = strVal1[1].trim(); }
			fnUpdateExcelCellData(strDataSheetPath, strDataSheet, strDataPointer, strColName, strVal); }
		else if (strDate.trim().toLowerCase().contains("+")) {
			String[] strDateVal = strDate.split("\\+");
			int strDateCnt = Integer.parseInt(strDateVal[1]);
			String strInputDate = fnGetObjectAttributeValue(strObj, "value").trim();
			if (strInputDate.trim().contains("|")){
				String[] strVal1 = strInputDate.split("\\|");
				strInputDate = strVal1[1].trim(); }
			String[] strTDate = strInputDate.split(" ");
			strVal = fnConvertDateFormat(strTDate[0], strTDate[1], strTDate[2], "dd/mm/yyyy");
			strVal = fnAddOrSubWorkingDays(strVal, strDateCnt, "add");
			strTDate = strVal.split("/");
			strVal = fnConvertDateFormat(strTDate[0], strTDate[1], strTDate[2], "dd mmm yyyy");
			fnUpdateExcelCellData(strDataSheetPath, strDataSheet, strDataPointer, strColName, strVal);	 }
		else if (strDate.trim().toLowerCase().contains("-")) {
			String[] strDateVal = strDate.split("\\-");
			int strDateCnt = Integer.parseInt(strDateVal[1]);
			String strInputDate = fnGetObjectAttributeValue(strObj, "value").trim();
			if (strInputDate.trim().contains("|")){
				String[] strVal1 = strInputDate.split("\\|");
				strInputDate = strVal1[1].trim(); }
			String[] strTDate = strInputDate.split(" ");
			strVal = fnConvertDateFormat(strTDate[0], strTDate[1], strTDate[2], "dd/mm/yyyy");
			strVal = fnAddOrSubWorkingDays(strVal, strDateCnt, "sub");
			strTDate = strVal.split("/");
			strVal = fnConvertDateFormat(strTDate[0], strTDate[1], strTDate[2], "dd mmm yyyy");
			fnUpdateExcelCellData(strDataSheetPath, strDataSheet, strDataPointer, strColName, strVal); }
		return strVal; }
	
	// function to enter the command deal through the command center ==================================================================================================
			public void fnEnterDealThroughCmdCenter(String strDataSheet, String strDataPointer, String product, String strFunctionality,String dealId) throws Throwable{
				//function call to clear the dash board
				fnClearDPMDashboard();

				//function call to clear the NOTIFICATIONS
				fnClearNotifications();
				String dealCommand = "";
				String tempFunctionality = "";

				if (strFunctionality.trim().toLowerCase().contains("unwind")) {
					tempFunctionality = "Unwind" ;
					//dealId = fnGetExcelData(pageObjects.strDataSheetPath, strDataSheet, strDataPointer, "DPM_Deal_ID");
					dealCommand = "/unwind " + dealId;
					String strTemp1 = fnGetExcelData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Unwind_Fee");
					String strTemp2 = fnGetExcelData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Unwind_FeeCurr");
					String strTemp3 = fnGetExcelData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Unwind_FeeSettlDate");
					String strTemp4 = fnGetExcelData(pageObjects.strDataSheetPath, dataSheet, dataPointer, "Unwind_FeeBook");
					dealCommand = "/unwind " + strTemp0 + " ";
					if (strTemp1.trim().length()>0){
						dealCommand = dealCommand + strTemp1 + " ";
						if (strTemp2.trim().length()>0){
							dealCommand = dealCommand + strTemp2 + " ";
							if (strTemp3.trim().length()>0){
								dealCommand = dealCommand + strTemp3 + " ";
								if (strTemp4.trim().length()>0){
									dealCommand = dealCommand + strTemp4 + " "; } } } } }
				else if(strFunctionality.trim().toLowerCase().contains("reverse")) {
					tempFunctionality = "Reverse" ;
					//dealId = fnGetExcelData(pageObjects.strDataSheetPath, strDataSheet, strDataPointer, "DPM_Deal_ID");
					dealCommand = "/reverse " + dealId; }

				System.out.println("[INFO]	Deal Command : "+dealCommand);


				if (fnObjectExists(objAddApps, 1)!=0) {
					fnObjectClick(objAddApps);
					fnAppSync(1);
					if (fnObjectExists(objCommandCenter, 1)!=0) {
						fnObjectClick(objCommandCenter);
						fnAppSync(1);
						if (fnObjectExists(objCommadcenterExpdBtn, 1)!=0)
							fnObjectClick(objCommadcenterExpdBtn);

						if (fnObjectExists(objCmdCntrLines, 2)!=0) {
							fnObjectDblClick(objCmdCntrLines);
							fnAppSync(1);

							//Step to enter the multiple command lines into the command center blotter
							Robot robot = new Robot();
							String strTemp1 = "";
							//dealCommand = "\\reverse E50362"+"\n\\reverse E50363"+"\n\\reverse E50364";
							if (dealCommand.trim().length()>0) {
								fnMouseMove(objCmdCntrLines);
								fnObjectClick(objCmdCntrLines);
								fnAppSync(1);
								fnObjectClick(objCmdCntrLines);
								int intCmdCnt = StringUtils.countMatches(dealCommand, ";");
								String[] strCmdList = dealCommand.split(";");
								for(int intLoop=0;intLoop<=intCmdCnt;intLoop++) {
									StringSelection strSelect = new StringSelection(strCmdList[intLoop].trim()+"\n");
									Toolkit.getDefaultToolkit().getSystemClipboard().setContents(strSelect, null);
									robot.keyPress(KeyEvent.VK_CONTROL);
									robot.keyPress(KeyEvent.VK_V);
									robot.keyRelease(KeyEvent.VK_V);
									robot.keyRelease(KeyEvent.VK_CONTROL);
									robot.keyRelease(KeyEvent.VK_ENTER);
									fnAppSync(1);
									strSelect = null; }
								//fnAppSync(1);
								robot.keyPress(KeyEvent.VK_BACK_SPACE);
								robot.keyRelease(KeyEvent.VK_BACK_SPACE); 

								pageObjects.strGVStepName = "Successfully Entered "+tempFunctionality+" Deal In Command Center - "+dealCommand;
								System.out.println("[INFO]   Successfully Entered "+tempFunctionality+" Deal In Command Center - "+dealCommand);
								fnScreenCapture(wDriver, strScreenPath.concat("dpmDealBookSuccess") + "-"+ strDataPointer);
							}

							String strTxt1="", strTxt2="", strTxt="";
							int strFound = -1;
							if (fnObjectExists(objCmdCntrDrpDwnMenu, 2)!=0) {
								fnObjectClick(objCmdCntrDrpDwnMenu);
								fnAppSync(1);
								if (fnObjectExists(objCmdCntrExecuteLnk, 1)!=0) {
									fnObjectClick(objCmdCntrExecuteLnk);
									strFound = -1;

									for(int intTemp = 0; intTemp<25; intTemp++) {
										fnAppSync(2);
										if (fnObjectExists(objBulkBkDealIDMsg, 1) != 0)
										{ 	Point objCoordinates = wDriver.findElement(By.xpath(objBulkBkDealIDMsg)).getLocation();
										robot.mouseMove(objCoordinates.getX()+63, objCoordinates.getY()+85);
										robot.mouseMove(objCoordinates.getX()+67, objCoordinates.getY()+81);
										strTxt1 = wDriver.findElement(By.xpath(objBulkBkDealIDMsg)).getText();

										System.out.println(">>>>>>>>> " + strTxt1);

										if (strTxt1.trim().toLowerCase().contains("object") == true) {
											strFound = 0;
											strTxt2 = "[object object]"; }
										else if (strTxt1.trim().toLowerCase().contains("invalid") == true)  {
											strFound = 0;
											strTxt2 = strTxt1; }
										else strFound = 1; }

										// ======================================================START======================================================================
										else if((fnObjectExists(objBulkBkDealIDMsg, 1) != 0) && (fnObjectExists(objBulkBkFailedToProcess, 1) != 0)){
											Point objCoordinates = wDriver.findElement(By.xpath(objBulkBkFailedToProcess)).getLocation();
											robot.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+87);
											//strTxt = wDriver.findElement(By.xpath(objBulkBkFailedToProcess)).getText();
											robot.mouseMove(objCoordinates.getX()+64, objCoordinates.getY()+83);
											strFound = 1;
										}
										else if(fnObjectExists(objBulkBkFailedToProcess, 1) != 0){
											Point objCoordinates = wDriver.findElement(By.xpath(objBulkBkFailedToProcess)).getLocation();
											robot.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+87);
											robot.mouseMove(objCoordinates.getX()+64, objCoordinates.getY()+83);

											if((fnObjectExists(objBulkBkDealIDMsg, 1) != 0) && (fnObjectExists(objBulkBkFailedToProcess, 1) != 0)){
												strFound = 1;
											}else{
												strTxt = wDriver.findElement(By.xpath(objBulkBkFailedToProcess)).getText();
												strFound = 0; pageObjects.strScnStatus = false;
											}
										}
										// ========================================================END====================================================================

										else if (fnObjectExists(objIDErrText, 1) != 0)
										{	Point objCoordinates = wDriver.findElement(By.xpath(objIDErrText)).getLocation();
										robot.mouseMove(objCoordinates.getX()+61, objCoordinates.getY()+87);
										strTxt = wDriver.findElement(By.xpath(objServerErrorContent)).getText();
										robot.mouseMove(objCoordinates.getX()+64, objCoordinates.getY()+83);
										strFound = 0; pageObjects.strScnStatus = false; }
										if (strFound != -1) break; }

									if ((strFound == 1) && (pageObjects.strScnStatus == true)) {
										boolean strSTxt = false;
										strTxt2 = wDriver.findElement(By.xpath(objBulkBkDealIDMsg)).getText();

										// ======================================================START======================================================================
										if(strTxt2.toLowerCase().contains("success") && strTxt2.toLowerCase().contains("fail")){
											strSTxt = true;
											fnGenerateHTMLRows(pageObjects.strHTMLPath, strDataPointer, "PASS&FAIL", strTxt2.substring(strTxt2.indexOf(":")+1, (strTxt2.indexOf("."))).trim(), pageObjects.strGVTime, strTxt2);
											strTxt2 = strTxt2.substring(strTxt2.indexOf(":")+1, (strTxt2.indexOf("."))).trim();
										}else if(strTxt2.toLowerCase().contains("success")){
											strSTxt = true;
											strTxt2 = strTxt2.substring(strTxt2.indexOf(":")+1, (strTxt2.trim().length() - 1)).trim();
										}
										// ========================================================END====================================================================

										if (strTxt2.toLowerCase().contains("Invalid lines:") == false) {
											if (fnObjectExists(objBulkBkDealIDMsg, 1) != 0)
												//strSTxt = (wDriver.findElement(By.xpath(objBulkBkDealIDMsg)).getText().toLowerCase().contains("success"));
												if (strSTxt == true) {
													pageObjects.strGVStepName = "Successfully Performed "+tempFunctionality+" functionality, Deal id - "+strTxt2;
													System.out.println("[INFO]   Successfully Performed "+tempFunctionality+" functionality, Deal id - "+strTxt2);
													fnScreenCapture(wDriver, strScreenPath.concat("dpmDealBookSuccess") + "-"+ strDataPointer); 
													fnUpdateExcelCellData(strDataSheetPath, strDataSheet, strDataPointer+"_"+tempFunctionality.trim(), "DPM_Deal_ID", strTxt2);
													String strTime = fnGetCurrentDate("yyyy/mm/dd") + " " + fnGetCurrentTime();
													fnUpdateExcelCellData(strDataSheetPath, strDataSheet, strDataPointer+"_"+tempFunctionality.trim(), "Deal_DateTime", strTime); } }
										else strFound = 0; }

									if ((strFound == 0 || strFound == -1) || (pageObjects.strScnStatus == false)) {
										if (fnObjectExists(objServerError, 1) != 0)
											strTxt = wDriver.findElement(By.xpath(objServerErrorContent)).getText();
										pageObjects.strScnStatus = false;
										pageObjects.strGVErrMsg = "Failed to "+tempFunctionality+" the " + pageObjects.strProductName + " bulk deal, error message : " + strTxt;
										pageObjects.strGVStepName = "Failed to "+tempFunctionality+" the " + pageObjects.strProductName + " bulk deal, error message : " + strTxt;
										fnScreenCapture(wDriver, strScreenPath.concat("dpmDealBookFailure") + "-"+ strDataPointer);
										fnGenerateHTMLRows(pageObjects.strHTMLPath, strDataPointer, "FAIL", pageObjects.strGVDealID, pageObjects.strGVTime, pageObjects.strGVErrMsg);
										System.out.println("[ERROR]   Failed to "+tempFunctionality+" the " + pageObjects.strProductName + " bulk deal, error message : " + strTxt); } } } }
						else { pageObjects.strScnStatus = false;
						System.out.println("[ERROR]   Failed to locate or input/edit in the Command Center");
						pageObjects.strGVStepName = "Failed to locate or input/edit in the Command Center";
						fnScreenCapture(wDriver, strScreenPath.concat("peakxvEQOCommandCenterBtnFailure") + "-" + strDataPointer);
						pageObjects.strGVErrMsg = "Failed to locate or input/edit in the Command Center"; }
					} else {
						pageObjects.strScnStatus = false;
						System.out.println("[ERROR]   Failed to find/ click the Command Center button");
						pageObjects.strGVStepName = "Failed to find/ click the Command Center button";
						fnScreenCapture(wDriver, strScreenPath.concat("peakxvEQOCommandCenterBtnFailure") + "-" + strDataPointer);
						pageObjects.strGVErrMsg = "Failed to find/ click the Command Center button"; }
				} else {
					pageObjects.strScnStatus = false;
					System.out.println("[ERROR]   Failed to find the ADD APPS button");
					pageObjects.strGVStepName = "Failed to find the ADD APPS button";
					fnScreenCapture(wDriver, strScreenPath.concat("peakxvEQOAddAppBtnFailure") + "-" + strDataPointer);
					pageObjects.strGVErrMsg = "Failed to find the ADD APPS button"; }}

		@SuppressWarnings({ "deprecation", "resource" })
		public List<String> fnGetColumnValues(String fileName, String sheetName, String colName) {
			String cellValue = "";
			int colNum = -1;
			List<String> columnValues = new ArrayList<String>();
			try{
			colNum = fnGetExcelColNum(fileName, sheetName, colName);
			}catch(Throwable t){
				System.out.println("[ERROR]	Exception Occured | comlib | fnGetColumnValues - "+t.getMessage());
			}
			if(colNum > -1 ){
				try{
					File file = new File(fileName);
					if (file.exists()) {
						InputStream myxls = new FileInputStream(new File(fileName));
						HSSFWorkbook book = new HSSFWorkbook(myxls);
						HSSFSheet sheet = book.getSheet(sheetName);
						Row row = sheet.getRow(0);

						for( int rowNum = 1 ; rowNum <= sheet.getLastRowNum() ; rowNum++){
							row = sheet.getRow(rowNum);
							Cell cell = row.getCell(colNum);
							if (cell == null)
								cellValue = "";
							else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
								cellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
							else if (cell.getCellType() == Cell.CELL_TYPE_STRING)
								cellValue = cell.getStringCellValue();
							else if (cell.getCellType() == Cell.CELL_TYPE_BLANK)
								cellValue = "";
							columnValues.add(cellValue);
						}
						myxls.close();
					}
				}catch(Exception e){
					System.out.println("[ERROR]	Exception Occured | comlib | fnGetColumnValues - "+e.getMessage());
				}
			}
			return columnValues;
		}

		@SuppressWarnings({ "deprecation", "resource" })
		public List<String> fnGetExcelRowValues(String fileName, String sheetName, int rowNum) {
			String cellValue = "";
			List<String> rowValues = new ArrayList<String>();
			if(rowNum > -1 ) {
				try{
					File file = new File(fileName);
					if (file.exists()) {
						InputStream myxls = new FileInputStream(new File(fileName));
						HSSFWorkbook book = new HSSFWorkbook(myxls);
						HSSFSheet sheet = book.getSheet(sheetName);
						Row row = sheet.getRow(rowNum);
						for( int cellNum = 0 ; cellNum < row.getLastCellNum() ; cellNum++){
							Cell cell = row.getCell(cellNum);
							if (cell == null)
								cellValue = "";
							else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
								cellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
							else if (cell.getCellType() == Cell.CELL_TYPE_STRING)
								cellValue = cell.getStringCellValue();
							else if (cell.getCellType() == Cell.CELL_TYPE_BLANK)
								cellValue = "";
							rowValues.add(cellValue);
						}
						myxls.close();
					}
				}catch(Exception e) {
					System.out.println("[ERROR]	Exception Occured | comlib | fnGetExcelRowValues - " + e.getMessage());	}	}
			return rowValues;	}
		
		//Function to enter the values into a array object by index reference ==========================================================================================
		public boolean fnEnterStringInObjectOfArrayByLocation(String strDataSheet, String strDataPointer, String strColName, String strObjXPATH, int intObjLoc) throws Throwable  {
			Robot robot = new Robot();
			boolean fnStatus = false;
			String strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strColName);
			if (strVal.isEmpty() == false){
				try{
					 List<WebElement> strList = wDriver.findElements(By.xpath(strObjXPATH)) ;
					 System.out.println("[INFO]    Input Object : " + strObjXPATH + " -- Count : " + strList.size() + " == Object Loc. " + intObjLoc);
					 if ((strList.size()-1) >= intObjLoc) {
						 strList.get(intObjLoc).clear();
						 strList.get(intObjLoc).sendKeys(strVal + Keys.ENTER);
						 fnAppSync(1);
						 strList.get(intObjLoc).sendKeys(Keys.TAB);
						 System.out.println("[INFO]    Successfully entered the details"); }
					 else
						 System.out.println("[ERROR]    Object list count is - 0");
				} catch (Exception exp){
					System.out.println("[ERROR]    Object not found/ failed to perform on - " + strObjXPATH); } }
			return fnStatus;	}
	
		//Function to enter the values into a array object by index reference ==========================================================================================
		public String fnGetStringInObjectOfArrayByLocation(String strObjXPATH, int intObjLoc, String strAttribute) throws Throwable  {
			Robot robot = new Robot();
			String stRetValue = "";
			try{
				 List<WebElement> strList = wDriver.findElements(By.xpath(strObjXPATH)) ;
				 System.out.println("[INFO]    Input Object : " + strObjXPATH + " -- Count : " + strList.size() + " == objLoc : " + intObjLoc);
				 if (strList.size() >= intObjLoc){
					 stRetValue = strList.get(intObjLoc).getAttribute(strAttribute);
					 System.out.println("[INFO]    Successfully retrived the details - " + stRetValue); }
				 else
					 System.out.println("[ERROR]    Object list count is - 0");
			} catch (Exception exp){
				System.out.println("[ERROR]    Object not found/ failed to perform on - " + strObjXPATH); }
			return stRetValue;	}

		//Function to click on a select object from an array object by index reference ==========================================================================================
		public boolean fnClickArrayObjByLocation(String strDataSheet, String strDataPointer, String strColName, String strObjXPATH, int intObjLoc) throws Throwable  {
			Robot robot = new Robot();
			boolean fnStatus = false;
			try{
				 List<WebElement> strList = wDriver.findElements(By.xpath(strObjXPATH)) ;
				 System.out.println("[INFO]    Click Object : " + strObjXPATH + " -- Count : " + strList.size());
				 if (strList.size() >= intObjLoc){
					 strList.get(intObjLoc).click();
					 System.out.println("[INFO]    Successfully entered the details"); }
				 else
					 System.out.println("[ERROR]    Object list count is - 0");
			} catch (Exception exp){
				System.out.println("[ERROR]    Object not found/ failed to perform on - " + strObjXPATH); }
			return fnStatus;	}

		//Function to enter the values into a array object by index reference ==========================================================================================
		public boolean fnSelectArrayObjectByIndexLocation(String strDataSheet, String strDataPointer, String strColName, String strObjXPATH, int intObjLoc, String strLower) throws Throwable  {
			Robot robot = new Robot();
			boolean fnStatus = false;
			String strVal = fnGetExcelData(strDataSheetPath, strDataSheet, strDataPointer, strColName);
			if (strLower.trim().toLowerCase().contentEquals("lower")) strVal = strVal.trim().toLowerCase();
			else if (strLower.trim().toLowerCase().contentEquals("upper")) strVal = strVal.trim().toUpperCase();
			if (strVal.isEmpty() == false){
				try {
					 List<WebElement> strList = wDriver.findElements(By.xpath(strObjXPATH));
					 System.out.println("[INFO]    List Object : " + strObjXPATH + " -- Count : " + strList.size() + " == intObjLoc : " + intObjLoc);
					 if (strList.size() >= intObjLoc){
						try{
							if (strList.get(intObjLoc).isEnabled()) {
								Select objSel = new Select(strList.get(intObjLoc));
								objSel.selectByValue(strVal);
								robot.delay(1000);
								System.out.println("[INFO]    Successfully select the value from drop down list"); } }
						catch(Exception exp){
							pageObjects.strScnStatus = false;
							pageObjects.strGVErrMsg = "Failed to search for value - " + strVal + "in the drop down selection";
							pageObjects.strGVStepName = "Failed to select/find " + strVal + " value from the dropdown";
							fnScreenCapture(wDriver, strScreenPath.concat("peakxvDealBookFailure") + "-"+ pageObjects.strGVScrDataPointer);
							pageObjects.strScnStatus = true;
							System.out.println("[ERROR]    ****** Failed to select the value : " + strVal + ", from the drop down object ID - " + strObjXPATH); } }
					 else
						 System.out.println("[ERROR]    Object list count is - 0"); }
				catch (Exception exp) {
					System.out.println("[ERROR]    Object not found/ failed to perform on - " + strObjXPATH); } }
			return fnStatus;	}

		//Function to enter the values into a array object by index reference ==========================================================================================
		public String fnGetArrayObjectByIndexLocation(String strObjXPATH, int intObjLoc, String strAttribute) throws Throwable  {
			Robot robot = new Robot();
			String strRetValue = "";
			try{
				List<WebElement> dropdown = wDriver.findElements(By.xpath(strObjXPATH));
				System.out.println("Object : " + strObjXPATH + " == Count : " + dropdown.size() + " == intObjLoc : " + intObjLoc);
				if (dropdown.size()>=intObjLoc){
					strRetValue = dropdown.get(intObjLoc).getAttribute(strAttribute); }
				else
					System.out.println("[ERROR]   Failed to read the proper no. of objects - " + strObjXPATH + " - Location : " + intObjLoc); }
			catch (Exception exp) {
				System.out.println("[ERROR]   Failed to read the value from object - " + strObjXPATH + " - Location : " + intObjLoc); }
			return strRetValue; }
		
//Function to pause the execution based on the config values ===================================================================================================================
		public void fnPauseExecStats() throws Throwable {
			Robot robot = new Robot();
			String strTempFilePath = pageObjects.strDefaultPath + pageObjects.strDefaultConfigName;
			String strValue = fnGetExcelData(strTempFilePath, "DPMConfig", "strPXVExecutionPauseStatus", "ConfigValue").trim();
			if (strValue.trim().length()==0) strValue = "execute";
			System.out.println("[INFO]   PXV Execution status - " + strValue + " - status .........");
			while(strValue.trim().toLowerCase().contentEquals("pause")) {
				System.out.println("[INFO]   PXV Execution status - " + strValue + " - waiting for status to change/update .........");
				robot.delay(10000);
				strValue = fnGetExcelData(strTempFilePath, "DPMConfig", "strPXVExecutionPauseStatus", "ConfigValue").trim(); } }

//Function which gives Json tagpath and json tagvalue as output ========================================================================================================
		public static String getJsonTokenValue(File jsonFile, String token) {
			String strRetVal = "";
			org.json.JSONObject jsonObject;
			try {
				jsonObject = getJSONObject(jsonFile.getAbsolutePath());
				String tokenToPreFix = "";
		
					if(null != token && token.contains(".")) {
						String[] tokenArr = token.split("\\.");
						tokenToPreFix = tokenArr[tokenArr.length - 1]; 
					} else {
						//return token +":"+ invokeTagsForPrinting(token, jsonObject);
						tokenToPreFix = token;
				}
					String tokenValue = invokeTagsForPrinting(token, jsonObject);
					if(null != tokenValue && !tokenValue.isEmpty()) {
						strRetVal = tokenToPreFix+":"+ invokeTagsForPrinting(token, jsonObject);
					}
				} catch (FileNotFoundException e) {
					e.printStackTrace();
				} catch (IOException e) {
					e.printStackTrace();
				}

			return strRetVal;		}

		public static String fnGetJsonTagData(String strFileName, String strJSONTag) throws Throwable {
			String strJsonFileTxt = fnReadFileMsg(strFileName);
			strJsonFileTxt = strJsonFileTxt.replace("\"", "").toLowerCase().trim();
			strJsonFileTxt = strJsonFileTxt.substring(1, strJsonFileTxt.length()-1);
			
			if (strJSONTag.trim().toLowerCase().contentEquals("trade.product.payLeg.swapLongFormSchedule.resetPeriodDate".toLowerCase()))
				System.out.println("break");
	
			char[] strJCHList = strJsonFileTxt.toCharArray();
			char strChar = '{';
			int cnt = 0;
			for(int intTemp=0;intTemp<strJCHList.length;intTemp++){
				if (strChar == strJCHList[intTemp])
					cnt = cnt + 1; }
			
			String strJTAGPath = strJSONTag.trim();
			String[] strList = strJTAGPath.trim().toLowerCase().split("\\.");

			int intPos = 0;
			String strTagFound  = "";
			int curTagPos = 0;
			int curTagPos1 = 0;
			for(int intTemp=0;intTemp<strList.length;intTemp++) {
				curTagPos = 0;
				curTagPos = strJsonFileTxt.indexOf(strList[intTemp]+":[{", intPos);

				if (curTagPos != -1) {
					if (strJSONTag.trim().toLowerCase().contentEquals("trade.product.payLeg.swapLongFormSchedule.resetPeriodDate".toLowerCase())==false
						&& strJSONTag.trim().toLowerCase().contentEquals("marketData.parcel.dimensions.surfacePoints.volatilityValue".toLowerCase())==false) {
						if (intTemp == strList.length-1) {
							intPos = curTagPos+4;
							curTagPos = strJsonFileTxt.indexOf(strList[intTemp]+":[", intPos);	} }
					else
						curTagPos = -1; 
				}

				if (curTagPos == -1) {
					if (intTemp < strList.length-1) {
						curTagPos = strJsonFileTxt.indexOf(strList[intTemp]+":{", intPos);
						if (curTagPos == -1)
							curTagPos = strJsonFileTxt.indexOf(strList[intTemp]+":[{", intPos);
						
						curTagPos1 = curTagPos + strList[intTemp].length() + 1; }
					else
						curTagPos = strJsonFileTxt.indexOf(strList[intTemp]+":", intPos);
					if (curTagPos != -1) {
						strTagFound = strTagFound + "." + strList[intTemp];	
						intPos = curTagPos + strList[intTemp].length();	}
					else if (curTagPos == -1) {
						strTagFound = "Not Found";
						break; } } }
			String st1 = "";
			if (strTagFound.trim().toLowerCase().contentEquals("not found")) {

			}
			else if (strTagFound.trim().toLowerCase().contentEquals("not found")==false) {
				char[] charArr = strJsonFileTxt.toCharArray();
				int intOpenCnt  = 0;
				int intOpenSCnt  = 0;
				int intIncCnt = 0;
				for(int intTemp=curTagPos1;intTemp<strJsonFileTxt.length();intTemp++) {
					if (charArr[intTemp]=='[') {
						intOpenSCnt = intOpenSCnt + 1;	}

					if (charArr[intTemp]=='{') {
						intOpenCnt = intOpenCnt + 1;
						intIncCnt = 1; }
					
					if (charArr[intTemp]=='}') {
						intOpenCnt = intOpenCnt - 1; }
					
					if (charArr[intTemp]==']') {
						intOpenSCnt = intOpenSCnt - 1;	}
					
					if (intOpenCnt == 1)
						st1 = st1 + charArr[intTemp];
					else if (intOpenSCnt == 1)
						st1 = st1 + charArr[intTemp];
					
					if (intOpenCnt == 0 && intIncCnt == 1 && intOpenSCnt == 0) {
						break;	} }
				
				int intComma=0;
				int intTem1 =0;
				if (st1.indexOf(st1.length()-1, 1)!='}')
					st1 = st1 + "}";
				String stTempColName = "";
				if (strJSONTag.trim().toLowerCase().contentEquals("trade.product.payLeg.swapLongFormSchedule.resetPeriodDate".toLowerCase()) 
						|| strJSONTag.trim().toLowerCase().contentEquals("trade.product.payLeg.swapLongFormSchedule.resetDate".toLowerCase())
						|| strJSONTag.trim().toLowerCase().contentEquals("trade.product.payLeg.swapLongFormSchedule.fxResetDate".toLowerCase())
						|| strJSONTag.trim().toLowerCase().contentEquals("trade.product.payLeg.swapLongFormSchedule.paymentPeriodDate".toLowerCase())
						|| strJSONTag.trim().toLowerCase().contentEquals("trade.product.payLeg.swapLongFormSchedule.paymentDate".toLowerCase())
						|| strJSONTag.trim().toLowerCase().contentEquals("trade.product.payLeg.swapLongFormSchedule.periodNotional".toLowerCase())
						|| strJSONTag.trim().toLowerCase().contentEquals("trade.product.payLeg.swapLongFormSchedule.spread".toLowerCase())
						|| strJSONTag.trim().toLowerCase().contentEquals("trade.product.payLeg.swapLongFormSchedule.interestPayment".toLowerCase())
						|| strJSONTag.trim().toLowerCase().contentEquals("trade.product.receiveLeg.swapLongFormSchedule.resetPeriodDate".toLowerCase()) 
						|| strJSONTag.trim().toLowerCase().contentEquals("trade.product.receiveLeg.swapLongFormSchedule.resetDate".toLowerCase())
						|| strJSONTag.trim().toLowerCase().contentEquals("trade.product.receiveLeg.swapLongFormSchedule.fxResetDate".toLowerCase())
						|| strJSONTag.trim().toLowerCase().contentEquals("trade.product.receiveLeg.swapLongFormSchedule.paymentPeriodDate".toLowerCase())
						|| strJSONTag.trim().toLowerCase().contentEquals("trade.product.receiveLeg.swapLongFormSchedule.paymentDate".toLowerCase())
						|| strJSONTag.trim().toLowerCase().contentEquals("trade.product.receiveLeg.swapLongFormSchedule.periodNotional".toLowerCase())
						|| strJSONTag.trim().toLowerCase().contentEquals("trade.product.receiveLeg.swapLongFormSchedule.spread".toLowerCase())
						|| strJSONTag.trim().toLowerCase().contentEquals("trade.product.receiveLeg.swapLongFormSchedule.interestPayment".toLowerCase())
						|| strJSONTag.trim().toLowerCase().contentEquals("trade.product.payLeg.swapLongFormSchedule.fixedRate".toLowerCase())
						|| strJSONTag.trim().toLowerCase().contentEquals("trade.product.receiveLeg.swapLongFormSchedule.fixedRate".toLowerCase())
						) {
					pageObjects.strArrList = "";
					String[] stListEle = st1.split(",");
					String[] stListEle1 = strJSONTag.split("\\.");
 					String stTagNm = stListEle1[stListEle1.length-1].trim().toLowerCase();
					for(int intTemp=0;intTemp<stListEle.length;intTemp++) {
						if (stListEle[intTemp].trim().toLowerCase().contains(stTagNm))
							pageObjects.strArrList  = pageObjects.strArrList + "#" + stListEle[intTemp]; 
						
					
					
					
					
					
					
					
					
					} }
				
				if (strTagFound.trim().toLowerCase().contains("holidaycenters")) {
					intTem1 = st1.indexOf("," + strList[strList.length-1]+":[", 0);
					if (intTem1 == -1)
						intTem1 = st1.indexOf("," + strList[strList.length-1]+":", 0);
					if (intTem1 == -1)
						intTem1 = st1.indexOf("{" + strList[strList.length-1]+":", 0);
					
					intComma = st1.indexOf("],", intTem1+strList[strList.length-1].length()+1);
					if (intComma == -1)
						intComma = st1.indexOf(",", intTem1+strList[strList.length-1].length()+1);
					if (intComma == -1)
						intComma = st1.indexOf("]}", intTem1+strList[strList.length-1].length()+1);	}
				else {
					intTem1 = st1.indexOf("," + strList[strList.length-1]+":", 0);
					if (intTem1 == -1)
						intTem1 = st1.indexOf("{" + strList[strList.length-1]+":", 0);
					
					intComma = st1.indexOf(",", intTem1 + strList[strList.length-1].length());	
					if (intComma == -1)
						intComma = st1.indexOf("},", intTem1 + strList[strList.length-1].length());
					if (intComma == -1)
						intComma = st1.indexOf("}",  intTem1 + strList[strList.length-1].length());	}

				if (intComma>0 && intTem1>=0) {
					int intTemp2 = intTem1 + strList[strList.length-1].trim().length() + 1;
					strJTAGPath = st1.substring(intTem1 , intComma);	}
				else
					strJTAGPath = strList[strList.length-1] + ":tag not found";	}
				strJTAGPath = strJTAGPath.replace("{", "").replace("}", "");
				if (strJTAGPath.substring(0,  1).contentEquals(",") || strJTAGPath.substring(0,  1).contentEquals("{"))
					strJTAGPath = strJTAGPath.substring(1,  strJTAGPath.length());
			return strJTAGPath;	}
		
		public static String fnReadFileMsg(String strFileName) throws Throwable {
	        File file = new File(strFileName);
	        if (file.exists()){
	             FileInputStream inputStream = null;
	             Scanner sc = null;
	             //Steps to read the parameters from config file
	          inputStream = new FileInputStream(strFileName);
	             sc = new Scanner(inputStream, "UTF-8");
	             String str1="";
	             while (sc.hasNextLine()) {
	                  String line = "";
	                  try {
	                       line = sc.nextLine().trim();
	                       str1 = str1 + line; }
	                  catch (Exception exp) {
	                       System.out.println("[ERROR]    Failed to read the line from the file - " + strFileName); } }
	             return str1;  }
	        else 
	             return "NO_FILE_EXISTS"; }
		

//=========================================================================================================================================================================
		static public org.json.JSONObject getJSONObject(String filePath) throws IOException, FileNotFoundException {
			JSONParser parser = new JSONParser();
			Object json_file_obj = null;
			try {
				json_file_obj = parser.parse(new FileReader(filePath));
			} catch (org.json.simple.parser.ParseException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			String str = json_file_obj.toString();
			org.json.JSONObject json_obj = new org.json.JSONObject(str);
			return json_obj;
		}

//=========================================================================================================================================================================
		static public String invokeTagsForPrinting(String tagName, org.json.JSONObject json_obj) {

			String[] strArr = tagName.split("\\.");
			org.json.JSONObject t_product = json_obj;
			String value = null;
			for(String tag : strArr){

				org.json.JSONObject temp = (org.json.JSONObject) t_product.optJSONObject(tag);
				if(temp != null){
					t_product = temp;
				} else {
					value = t_product.optString(tag);
					if(value != null && value != "" && !value.equals("") ) {
						System.out.println(tagName + ": " + value);
						return value;
					}
				}
			}

			System.out.println("---------------------------------------------------------------------------------------------------");
			System.out.println("Printing values of: " + tagName);
			System.out.println("---------------------------------------------------------------------------------------------------");
			for (String o : t_product.keySet()) {
				Object object = t_product.get(o);
				if (object instanceof org.json.JSONObject) {
					//System.out.println("Printing sub-elements of: " + o);
					printJSONObject(object);
				} else if (object instanceof JSONArray) {
					//System.out.println("json array: Key:" + o + "; Value:" + t_product.get(o));
				} else {
					//System.out.println(o + ": " + t_product.get(o));
					return t_product.get(o).toString();
				}
			}
			return value;
		}

//=========================================================================================================================================================================
		static public String printJSONObject(Object o) {
			org.json.JSONObject json_obj = (org.json.JSONObject) o;
			Object value = null;
			for (String obj_str : json_obj.keySet()) {
				Object object = json_obj.get(obj_str);
				if (object instanceof org.json.JSONObject) {
					//System.out.println("     Printing sub-elements of: " + obj_str);
					printJSONObject(object);
				} else if (object instanceof JSONArray) {
					value = object; 
					//System.out.println("json array: Key:" + obj_str + "; Value:" + json_obj.get(obj_str));
				} else {
					value =  json_obj.get(obj_str);
					//System.out.println(obj_str + ": " + json_obj.get(obj_str));
				}
			}
			if (null != value) {
				return value.toString();
			} else {
				return null;
			}
		}

		

		
		//Function which gives Json tagpath and json tagvalue as output
				public static String getJsonTokenValue(File jsonFile, String prod, String token) {
					
					org.json.JSONObject jsonObject;
					try {
						jsonObject = getJSONObject(jsonFile.getAbsolutePath());
					
						File mapping_file = null;
					if ("fra".equalsIgnoreCase(prod)) {
						mapping_file = new File("\\\\172.23.43.24\\dpm-qa\\Docs\\PXV_Automation_TestData\\XVA_Deals_Validate\\FRA_DataFields_Mapping.txt");
					} else if ("Forex".equalsIgnoreCase((prod))) {
						mapping_file = new File("\\\\172.23.43.24\\dpm-qa\\Docs\\PXV_Automation_TestData\\XVA_Deals_Validate\\FXFORWARD_DataFields_Mapping.txt");
					} else if ("interestrates".equalsIgnoreCase(prod)) {
						mapping_file = new File("\\\\172.23.43.24\\dpm-qa\\Docs\\PXV_Automation_TestData\\XVA_Deals_Validate\\InterestRates_DataFields_Mapping.txt");
					} else if ("FX: Spot Forward".equalsIgnoreCase(prod)) {
						mapping_file = new File("\\\\172.23.43.24\\dpm-qa\\Docs\\PXV_Automation_TestData\\XVA_Deals_Validate\\MUREX_WSS_DataFields_Mapping.txt");
					} else if("CC Swap".equalsIgnoreCase((prod))) {
						mapping_file=new File("\\\\172.23.43.24\\dpm-qa\\Docs\\PXV_Automation_TestData\\XVA_Deals_Validate\\SC_CC_SWAP_DataFields_Mapping.txt");
					}
					if (null != prod) {
						FileReader mappingFileReader = new FileReader(mapping_file);
						BufferedReader mappingBufferedReader = new BufferedReader(mappingFileReader);
						String line = null;
						while ((line = mappingBufferedReader.readLine()) != null) {
							String tagArr[] = line.split(":");
							String[] mappingKey = tagArr[0].split("\\.");
							
							if(tagArr[0].equalsIgnoreCase(token)) {
								return mappingKey[mappingKey.length - 1] +":"+ invokeTagsForPrinting(tagArr[0], jsonObject);
							}
						}
						mappingFileReader.close();
					} else {
						String[] mappingKey = token.split("\\.");
						return mappingKey[mappingKey.length - 1] +":"+ invokeTagsForPrinting(token, jsonObject);
					}
					
						} catch (FileNotFoundException e) {
							//
							e.printStackTrace();
						} catch (IOException e) {
							// 
							e.printStackTrace();
						}
					return null;
				}
				
				static public org.json.JSONObject getJSONObject(String filePath) throws IOException, FileNotFoundException {
					JSONParser parser = new JSONParser();
					Object json_file_obj = null;
					try {
						json_file_obj = parser.parse(new FileReader(filePath));
					} catch (org.json.simple.parser.ParseException e) {
						// 
						e.printStackTrace();
					}
					String str = json_file_obj.toString();
					org.json.JSONObject json_obj = new org.json.JSONObject(str);
					return json_obj;
				}

				
				static public String invokeTagsForPrinting(String tagName, org.json.JSONObject json_obj) {

					String[] strArr = tagName.split("\\.");
					org.json.JSONObject t_product = json_obj;
					String value = null;
					for(String tag : strArr){

						org.json.JSONObject temp = (org.json.JSONObject) t_product.optJSONObject(tag);
						if(temp != null){
							t_product = temp;
						} else {
							value = t_product.optString(tag);
							if(value != null && value != "" && !value.equals("") ) {
								System.out.println(tagName + ": " + value);
								return value;
							}
						}
					}

					System.out.println("---------------------------------------------------------------------------------------------------");
					System.out.println("Printing values of: " + tagName);
					System.out.println("---------------------------------------------------------------------------------------------------");
					for (String o : t_product.keySet()) {
						Object object = t_product.get(o);
						if (object instanceof org.json.JSONObject) {
							System.out.println("Printing sub-elements of: " + o);
							printJSONObject(object);
						} else if (object instanceof JSONArray) {
							System.out.println("json array: Key:" + o + "; Value:" + t_product.get(o));
						} else {
							System.out.println(o + ": " + t_product.get(o));
							return t_product.get(o).toString();
						}

					}
					return value;
				}

				static public String printJSONObject(Object o) {
					org.json.JSONObject json_obj = (org.json.JSONObject) o;
					Object value = null;
					for (String obj_str : json_obj.keySet()) {
						Object object = json_obj.get(obj_str);
						if (object instanceof org.json.JSONObject) {
							System.out.println("     Printing sub-elements of: " + obj_str);
							printJSONObject(object);
						} else if (object instanceof JSONArray) {
							value = object; 
							System.out.println("json array: Key:" + obj_str + "; Value:" + json_obj.get(obj_str));
						} else {
							value =  json_obj.get(obj_str);
							System.out.println(obj_str + ": " + json_obj.get(obj_str));
						}
					}
					return value.toString();
				}

				
		
		
		
}*/