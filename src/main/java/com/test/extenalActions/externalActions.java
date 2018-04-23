package com.test.extenalActions;
import java.awt.Robot;
import java.awt.event.KeyEvent;

import org.junit.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;
import org.testng.Reporter;
import com.test.testResources.*;

public class externalActions extends pageObjects {
	static commnFunctions fnCommLib = new commnFunctions();
	
	//Method to select the specified product from the menu ========================================================================
	public boolean fnSelectMenuSubMenu(String strDataSheet, String strDataPointer, String strColName) throws Throwable {
		boolean strRetVal = false;

		//Step to search for product category link ================================================================================
		if (fnCommLib.fnObjectExists(btnProducts, 1)!=0) {
			log.error("Found Product Category link - " + btnProducts);
			Reporter.log("Found Product Category link");
			fnCommLib.fnMouseMove(btnProducts);
			
			//Step to search for product category sub-menu link ===================================================================
			String strValue = fnCommLib.fnGetExcelData(pageObjects.strDefaultPath, strDataSheet, strDataPointer, strColName).trim();
			String strPath = pageObjects.btnProductsSubItem.replace("VALUE", strValue);
			log.info("Searching for menu option : " + strPath);
			
			if (fnCommLib.fnObjectExists(strPath, 1)!=0) {
				log.info("Found " + strValue + " link : " + strPath);
				Reporter.log("Found " + strValue + " link");
				fnCommLib.fnCaptureScreenShot("Found " + strValue + " link");

				fnCommLib.fnObjectClick(strPath);
				fnCommLib.fnAppSync(1);
				strRetVal = true; }
			else {
				stepStatus = false;
				log.error("Failed to search " + strValue + " link : " + strPath);
				Reporter.log("Failed to search " + strValue + " link");
				testAssert.fail("Failed to search " + strValue + " link");
				Assert.fail("Failed to search " + strValue + " link");
				fnCommLib.fnCaptureScreenShot("Failed to search " + strValue + " link");
				} }
		else {
			stepStatus = false;
			log.error("Failed to search Product Category link : " + btnProducts);
			Reporter.log("Failed to search Product Category link");
			testAssert.fail("Failed to search Product Category link");
			Assert.fail("Failed to search Product Category link");
			fnCommLib.fnCaptureScreenShot("Failed to search Product Category link"); }
		if (strRetVal == false) pageObjects.stepStatus = false;
		return strRetVal;	}

	//Method to select the specified product from the menu ========================================================================
	public static boolean fnValidateProductTransactionStatus() throws Throwable {
		boolean strRetVal = false;
		if (fnCommLib.fnObjectExists(lblTransactionResults, 2)!=0) {
			fnCommLib.fnObjectClick(lblTransactionResults);
			log.info("Successfully completed the transaction/order");
			Reporter.log("Successfully completed the transaction/order");
			fnCommLib.fnCaptureScreenShot("Successfully completed the transaction/order");
			strRetVal = true;}
		else if (fnCommLib.fnObjectExists(msgError, 2)!=0) {
			fnCommLib.fnObjectClick(msgError);
			String strErr = fnCommLib.fnGetObjectAttributeValue(msgError, "outerText");
			log.error("Failed to place the transaction/order, Error : " + strErr);
			Reporter.log("Failed to place the transaction/order, Error : " + strErr);
			testAssert.fail("Failed to place the transaction/order, Error : " + strErr);
			stepStatus = false;
			fnCommLib.fnCaptureScreenShot("Failed to place the transaction/order, Error : " + strErr); }
		else {
			log.error("Failed to place the transaction/order");
			Reporter.log("Failed to place the transaction/order");
			testAssert.fail("Failed to place the transaction/order");
			stepStatus = false;
			fnCommLib.fnCaptureScreenShot("Failed to place the transaction/order"); }
		return strRetVal;	}

	//Method to select the specified product from the menu ========================================================================
	public boolean fnAddToCartProduct(String strDataSheet, String strDataPointer, String strColName) throws Throwable {
		boolean strRetVal = false;
		//Step to search for product category link ================================================================================
		String strValue = fnCommLib.fnGetExcelData(pageObjects.strDefaultPath, strDataSheet, strDataPointer, strColName).trim();
		String stPath = lnkProductTxt.replace("ITEM-TEXT", strValue);

		if (fnCommLib.fnObjectExists(stPath, 1)!=0) {
			log.info("Found Product Item link - " + stPath);
			Reporter.log("Found Product - " + strValue);
		
			stPath = lnkItemAddToCart.replace("ITEM-TEXT", strValue);
			
			if (fnCommLib.fnObjectExists(stPath, 1)!=0) {
				fnCommLib.fnObjectClick(stPath);
				log.info("Add to cart operation successful - " + stPath);
				Reporter.log("Add to cart operation successful");
				fnCommLib.fnCaptureScreenShot("Add to cart operation successful");
				strRetVal = true; }
			else {
				log.error("Add to cart operation failed - " + stPath);
				Reporter.log("Add to cart operation failed - " + strValue);
				testAssert.fail("Add to cart operation failed - " + strValue);
				stepStatus = false;
				fnCommLib.fnCaptureScreenShot("Add to cart operation failed - " + strValue); }	}
		else {
			log.error("Failed to locate Product Item - " + stPath);
			Reporter.log("Failed to search for product - " + strValue);
			testAssert.fail("Failed to search for product - " + strValue);
			stepStatus = false;
			fnCommLib.fnCaptureScreenShot("Failed to search for product - " + strValue); }
		if (strRetVal == false) pageObjects.stepStatus = false;
		return strRetVal;	}

	//Method to validate the add cart value ========================================================================
	public boolean fnAddToCartProductVerification(String strDataSheet, String strDataPointer, String strColName) throws Throwable {
		boolean strRetVal = false;
		String strValue="0";
		if (strColName.trim().isEmpty()==false) {
			strValue = fnCommLib.fnGetExcelData(pageObjects.strDefaultPath, strDataSheet, strDataPointer, strColName).trim();
			if (strValue.trim().isEmpty()==true) {
				strValue = "0"; } }
		//Step to verify the add cart count value object exists for not ==============================================
		if (fnCommLib.fnObjectExists(txtAddCartCount, 1)!=0) {
			String intActVal = fnCommLib.fnGetObjectAttributeValue(txtAddCartCount, "outerText");
			if (intActVal.trim().contentEquals(strValue.trim())) {
				log.info("Cart item count validation successful - Expected : " + strValue + " == Actual : " + intActVal);
				Reporter.log("Cart item count validation successful - Expected : " + strValue + " == Actual : " + intActVal);
				fnCommLib.fnCaptureScreenShot("Cart item count validation successful - Expected : " + strValue + " == Actual : " + intActVal);
				strRetVal = true; }
			else {
				log.error("Cart item count validation failed - Expected : " + strValue + " == Actual : " + intActVal);
				Reporter.log("Cart item count validation failed - Expected : " + strValue + " == Actual : " + intActVal);
				testAssert.fail("Cart item count validation failed - Expected : " + strValue + " == Actual : " + intActVal);
				stepStatus = false;
				fnCommLib.fnCaptureScreenShot("Cart item count validation failed - Expected : " + strValue + " == Actual : " + intActVal); }
		} else {
			log.error("Failed to locate Add card items count - " + txtAddCartCount);
			Reporter.log("Failed to locate Add card items count - " + strValue);
			testAssert.fail("Failed to locate Add card items count - " + strValue);
			stepStatus = false;
			fnCommLib.fnCaptureScreenShot("Failed to locate Add card items count - " + strValue); }
		if (strRetVal == false) pageObjects.stepStatus = false;
		return strRetVal;	}
	
	//Method to validate the add cart value ===========================================================================
	public static boolean fnEnterProductAddressDetils(String strDataSheet, String strDataPointer) throws Throwable {
		boolean strRetVal = false;
		Robot robot = new Robot();

		//Steps to set the billing address details ==========================================================================
		strRetVal =  fnCommLib.fnObjectSendKeys(txtEmail, "Email Address", strDataSheet, strDataPointer, "EmailAddress");
		
		strRetVal =  fnCommLib.fnObjectSendKeys(txtFirstName, "First Name", strDataSheet, strDataPointer, "FirstName");
		
		strRetVal =  fnCommLib.fnObjectSendKeys(txtLastName, "Last Name", strDataSheet, strDataPointer, "LastName");
		
		strRetVal =  fnCommLib.fnObjectSendKeys(txtAddress, "Address", strDataSheet, strDataPointer, "Address");
		
		strRetVal =  fnCommLib.fnObjectSendKeys(txtCity, "City", strDataSheet, strDataPointer, "City");
		
		strRetVal =  fnCommLib.fnObjectSendKeys(txtState, "State", strDataSheet, strDataPointer, "State");

		robot.keyPress(KeyEvent.VK_PAGE_DOWN);
		robot.keyRelease(KeyEvent.VK_PAGE_DOWN);
		strRetVal =  fnCommLib.fnSelectByKeys(lblBillCountry, "Country", strDataSheet, strDataPointer, "Country");
		
		strRetVal =  fnCommLib.fnObjectSendKeys(txtPostalCode, "Postal Code", strDataSheet, strDataPointer, "PostalCode");
		
		strRetVal =  fnCommLib.fnObjectSendKeys(txtPhone, "Phone", strDataSheet, strDataPointer, "Phone");
		
		//Steps to set the shipping address details =========================================================================

		strRetVal =  fnCommLib.fnObjectCheck(chkSameAsBilling, "Same as billing address", strDataSheet, strDataPointer, "SameAsBillingAddress");
		
		WebElement ele = wDriver.findElement(By.xpath(chkSameAsBilling));
		if (ele.isSelected()==false) {
			strRetVal =  fnCommLib.fnObjectSendKeys(txtShipFirstName, "First Name", strDataSheet, strDataPointer, "SA_FirstName");
			
			strRetVal =  fnCommLib.fnObjectSendKeys(txtShipLastName, "Last Name", strDataSheet, strDataPointer, "SA_LastName");
			
			strRetVal =  fnCommLib.fnObjectSendKeys(txtShipAddress, "Address", strDataSheet, strDataPointer, "SA_Address");
			
			strRetVal =  fnCommLib.fnObjectSendKeys(txtShipCity, "City", strDataSheet, strDataPointer, "SA_City");
			
			strRetVal =  fnCommLib.fnObjectSendKeys(txtShipState, "State", strDataSheet, strDataPointer, "SA_State");
	
			robot.keyPress(KeyEvent.VK_PAGE_DOWN);
			robot.keyRelease(KeyEvent.VK_PAGE_DOWN);
			strRetVal =  fnCommLib.fnSelectByKeys(lblShipCountry, "Country", strDataSheet, strDataPointer, "SA_Country");
			
			strRetVal =  fnCommLib.fnObjectSendKeys(txtShipPostalCode, "Postal Code", strDataSheet, strDataPointer, "SA_PostalCode"); }
		
		fnCommLib.fnCaptureScreenShot("Successfully entered the Billing & Shipping address details");
		return strRetVal; }

}
