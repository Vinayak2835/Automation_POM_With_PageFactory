package com.pom.test;

import org.testng.Assert;
import org.testng.annotations.Test;
import com.pom.base.BaseClass;
import com.pom.pages.HomePage;
import com.pom.pages.LoginPage;

public class HomePageTest extends BaseClass {
	
	HomePage homePage;
	//LoginPage loginPage;
	
	@Test(priority = 3)
	public void adminTagTest( ) throws InterruptedException {
		
		homePage = loginPage.login("Admin","admin123");
		homePage.clickOnAdminTab();
		Thread.sleep(2000);
		
		String expectedUrl = "https://opensource-demo.orangehrmlive.com/index.php/admin/viewSystemUsers";
		String actualUrl = driver.getCurrentUrl();
		Assert.assertEquals( actualUrl,expectedUrl);
	}

}
