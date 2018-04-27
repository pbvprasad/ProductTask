package com.test.testResources;
import java.awt.Robot;
import java.io.PrintStream;

import org.apache.log4j.Logger;
import org.openqa.selenium.WebDriver;
import org.testng.asserts.SoftAssert;
public class pageObjects {

	// global variables ======================================================================================================
	public static String strDefaultPath = "C:\\Ethoca_Task\\TestData\\ethoca_ProductOrderCases_TestData.xls";
	public static String strChromeDriverPathName = "C:\\ChromeDriver\\chromedriver.exe";
	
	public static String strDefaultChromeFileName = "chromedriver.exe";
	public static WebDriver wDriver = null, wDriver1=null, wDriver2=null;
	public static String strFolderPath;
	public static String strHTMLName;
	public static int intUpdateWaitTime = 2;
	public static int intCaseCnt = 0;
	public static String strHTMLPath = "C:\\Ethoca_Screens\\";
	public static PrintStream stdout = System.out;
	public static String strDataSheetPath;
	public static String strResultsPath;
	public static String strConfigFilepath;
	public static String strConfigFileName = "peakxv_Configuration.xls";
	public static String strDataFileName = null;
	public static String strDataFilePath = null;
	public static String strProductName;
	public static String strGVStepName = "";
	
	public static boolean strDefaultValues = false;
	public static boolean strScreenCapture = true;
	public static boolean strScrnFileName = false;
	public static boolean strGenPDF = true;
	public static String strScreenPath = "";
	public static String strPDFPath = "";
	public static String strScreenPath1 = "";
	public static int intSheetRowNum = 0;
	public static int intImgCnt  = 0;
	public static int strScreenFileCnt = 0;
	public static String strEnvironmentURL = "";
	public static String strImgList[][] = new String[1][];
	public static String strStepList[][] = new String[1][]; 
	public static String strStepListStatus[][] = new String[1][];
	public static boolean strScnStatus = true;
	public static String strGVDate = "";
	public static String strGVTime = "";
	public static SoftAssert testAssert = new SoftAssert();

	public String log4jConfigFile = "/Ethoca_Task/log4j.properties";
	public static Logger log = null;
	public static boolean stepStatus = true;

	//PAGE OBJECTS ==================================================================================================================	
	public static String lblToolsImg = "//header[@id='header']//a[@id='logo']";
	public static String btnProducts = "//ul[@id='menu-main-menu' and @class='menu']//a[contains(text(),'Product Category')]";
	public static String btnProductsSubItem = "//ul[@class='sub-menu']//a[contains(text(),'VALUE')]";
	public static String lnkProductTxt = "//a[contains(text(),'ITEM-TEXT')]";
	public static String lnkItemAddToCart = "//a[contains(text(),'ITEM-TEXT')]/../..//input[@name='Buy']";
	public static String txtAddCartCount = "//div[@id='header_cart']//em[@class='count']";
	public static String lnkCheckOut = "//div[@id='header_cart']//span[@class='icon' and contains(text(),'Cart')]";
	public static String lblCheckOut = "//header[@class='page-header']//h1[contains(text(),'Checkout')]";
	public static String btnCheckOutContinue = "//div[@id='checkout_page_container']//span[contains(text(),'Continue')]";
	public static String btnPurchase = "//div[@class='wpsc_make_purchase']//input[@value='Purchase']";
	public static String lstCountry = "//table[@class='productcart']//select[@class='current_country wpsc-visitor-meta wpsc-country-dropdown']";

	public static String lblCountry = "//div[@id='uniform-current_country']";
	
	public static String txtEmail = "//label[contains(text(),'Enter your email address')]/input";
	public static String txtFirstName = "//input[@title='billingfirstname']";
	public static String txtLastName = "//input[@title='billinglastname']";
	public static String txtAddress = "//textarea[@title='billingaddress']";
	public static String txtCity = "//input[@title='billingcity']";
	public static String txtState = "//input[@title='billingstate']";
	public static String lblBillCountry = "//div[@id='uniform-wpsc_checkout_form_7']";
	public static String txtCountry = "//select[@title='billingcountry']";
	public static String txtPostalCode = "//input[@title='billingpostcode']";
	public static String txtPhone = "//input[@title='billingphone']";
	
	public static String chkSameAsBilling = "//label[contains(text(),'Same as billing address')]/../..//input";
	public static String txtShipFirstName = "//input[@title='shippingfirstname']";
	public static String txtShipLastName = "//input[@title='shippinglastname']";
	public static String txtShipAddress = "//textarea[@title='shippingaddress']";
	public static String txtShipCity = "//input[@title='shippingcity']";
	public static String txtShipState = "//label[contains(text(),'undefined')]/../..//input[@title='shippingstate']";
	public static String lblShipCountry = "//div[@id='uniform-wpsc_checkout_form_16']";
	public static String txtShipCountry = "//select[@title='shippingcountry']";
	public static String txtShipPostalCode = "//input[@title='shippingpostcode']";
	
	public static String btnConfirmPurchase = "//input[@value='Purchase' and @type='submit']";
	public static String lblTransactionResults = "//h1[contains(text(),'Transaction Results')]";
	public static String msgError = "//div[@class='login_error']/p[@class='validation-error']";
	
	
	
	
	
	
	
	
	
	
	
	
	
	




}