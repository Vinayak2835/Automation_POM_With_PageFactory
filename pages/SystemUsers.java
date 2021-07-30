package com.pom.pages;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;

public class SystemUsers {
	@FindBy(xpath="//*[@id=\"menu_admin_viewAdminModule\"]/b")
	WebElement adminTag;
	
	public void clickOnAdminTab() {
		adminTag.click();
	}
}
