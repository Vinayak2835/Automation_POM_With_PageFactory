package com.pom.test;

import org.testng.Assert;
import org.testng.annotations.Test;

import com.pom.base.BaseClass;
import com.pom.pages.HomePage;
import com.pom.pages.LoginPage;

public class LoginPageTest extends BaseClass {
	
	HomePage homePage;
	
	@Test(priority = 1)
	public void verifyLogo() {
		boolean result = loginPage.validateLogo();
		Assert.assertTrue(result);		
	}
	
	@Test(priority = 2)
	public void loginTest() {
		homePage = loginPage.login("Admin","admin123");
		String expectedUrl = "https://opensource-demo.orangehrmlive.com/index.php/dashboard";
		String actualUrl = driver.getCurrentUrl();
		Assert.assertEquals( actualUrl,expectedUrl);
	}

}
