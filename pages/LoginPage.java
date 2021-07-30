package com.pom.pages;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import com.pom.base.BaseClass;

public class LoginPage extends BaseClass {
	
	@FindBy(id="txtUsername") 
	WebElement userName;
	
	@FindBy(name= "txtPassword") 
	WebElement userPassword;
	
	@FindBy(id="btnLogin")
	WebElement loginBtn;
	
	@FindBy(xpath="//*[@id=\"divLogo\"]/img")
	WebElement logo;
	
	public LoginPage() {
		PageFactory.initElements(driver, this);
	}
	
	public boolean validateLogo() {
		logo.isDisplayed();
		
		return true;
	}
	
	public HomePage login(String uname, String password) {
		userName.sendKeys(uname);
		userPassword.sendKeys(password);
		loginBtn.click();
		
		return new HomePage();
	}

}
