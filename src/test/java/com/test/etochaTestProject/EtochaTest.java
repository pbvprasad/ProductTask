package com.test.etochaTestProject;
import org.openqa.selenium.By;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.Reporter;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import com.test.extenalActions.externalActions;
import com.test.testResources.commnFunctions;
import com.test.testResources.pageObjects;

public class EtochaTest extends pageObjects {
	commnFunctions fnCommLib = new commnFunctions();
	externalActions testSteps = new externalActions();

	@BeforeTest
	public void beforeTest() throws Throwable {
		//Steps to load the confirmation parameters if any ===========================================
		fnCommLib.fnLoadConfigurationParameters(); }

	@Test(priority=0)
	public void productOrder_Validation() throws Throwable {
		String strDataSheet = "Products_BuyTestCases";
		pageObjects.stepStatus = true;
		pageObjects.intImgCnt = 1;
		pageObjects.strImgList[0] = new String[25];
		pageObjects.strStepList[0] = new String[25];
		pageObjects.strStepListStatus[0] = new String[25];
		int intRowCount  = fnCommLib.fnGetExcelRowCount(pageObjects.strDataSheetPath, strDataSheet);

		//Steps to launch the browser ================================================================
		fnCommLib.dpmChromeBrowserLaunch("Products_BuyTestCases", "Order_TC_1");
		
		fnCommLib.fnGenerateHTMLHeaderReport(pageObjects.strHTMLPath);
		
		//Step to click on the menu link =============================================================
		if (pageObjects.stepStatus == true) {
			log.info("Validating the page loaded successfully or not - " + wDriver.getTitle());
			if (wDriver.getTitle().trim().toLowerCase().contains("Toolsqa Dummy Test site".toLowerCase())) {
				log.info("Successfully opened the AUT");
				Reporter.log("Successfully opened the AUT"); 
				fnCommLib.fnCaptureScreenShot("Successfully opened the AUT");  }
			else if (wDriver.getTitle().trim().toLowerCase().contains("Toolsqa Dummy Test site".toLowerCase())==false) {
				stepStatus = false;
				log.error("Failed to open the AUT");
				Reporter.log("Failed to open the AUT");
				Assert.fail("Failed to open the AUT");
				testAssert.fail("Failed to open the AUT");
				fnCommLib.fnCaptureScreenShot("Failed to open the AUT"); } }
		
		if (pageObjects.stepStatus == true)
			stepStatus = testSteps.fnSelectMenuSubMenu(strDataSheet, "Order_TC_1", "ProductType");
		
		if (pageObjects.stepStatus == true)
			stepStatus = testSteps.fnAddToCartProduct(strDataSheet, "Order_TC_1", "ProductName");
		
		if (pageObjects.stepStatus == true) {
			WebDriverWait wait=new WebDriverWait(wDriver, 20);
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(btnProducts)));
			
			fnCommLib.fnAppSync(10);
			stepStatus = testSteps.fnAddToCartProductVerification(strDataSheet, "Order_TC_1", "ProductQuantity"); }

		//Step to validate the checkout button, click and navigate to the check out page================================
		if (pageObjects.stepStatus == true) {
			stepStatus = fnCommLib.fnValidatePageNavigationStatus(lnkCheckOut, "Check Out", lblCheckOut); }

		//Step to navigate to the address/ billing page ================================================================
		if (pageObjects.stepStatus == true) {
			stepStatus = fnCommLib.fnValidatePageNavigationStatus(btnCheckOutContinue, "Purchase", btnPurchase); }
		
		//Step to enter the details in the address/ billing page ================================================================
		if (pageObjects.stepStatus == true) {
			stepStatus = externalActions.fnEnterProductAddressDetils(strDataSheet,  "Order_TC_1");	}
	
		//Steps to validate the transaction success page =====================================================================================
		if (pageObjects.stepStatus == true) {
			fnCommLib.fnObjectClick(btnConfirmPurchase);
			log.info("Successfully click on the Purchase button");
			Reporter.log("Successfully click on the Purchase button");
			stepStatus = externalActions.fnValidateProductTransactionStatus(); }
		
		//step to create the PDF document with the list of screen shots
		fnCommLib.fnPDFScreenReport("Order_TC_1");
		if (pageObjects.stepStatus == true)
			fnCommLib.fnGenerateHTMLRows(pageObjects.strHTMLPath, "Order_TC_1", "PASS", "TIME", pageObjects.strGVTime, "Successfully completed the product deal");	
		else
			fnCommLib.fnGenerateHTMLRows(pageObjects.strHTMLPath, "Order_TC_1", "FAIL", "TIME", pageObjects.strGVTime, "Failed to book the product deal");
		testAssert.assertAll();	}

	@AfterTest
	public void afterTest() throws Throwable {
		log.info("Closing the browsers");
		if (wDriver != null) {
			wDriver.close();
			Thread.sleep(3000);
			wDriver.quit();
			}
		wDriver = null;	
		}
}
