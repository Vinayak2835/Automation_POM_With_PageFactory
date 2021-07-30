package com.pom.pages;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import com.pom.base.BaseClass;

public class HomePage extends BaseClass {
	
	@FindBy(xpath="//*[@id=\"menu_admin_viewAdminModule\"]/b")
	WebElement adminTab;
	
	public HomePage() {
		PageFactory.initElements(driver, this);
	}
	
	public SystemUsers clickOnAdminTab() {
		adminTab.click();
		return new SystemUsers();
	}
}
