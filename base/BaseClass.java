package com.pom.base;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import com.pom.pages.LoginPage;
import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {
    public static WebDriver driver;
    public LoginPage loginPage;
  
  @BeforeMethod
  public void setup() {
	
	     WebDriverManager.chromedriver().setup();
		 driver = new ChromeDriver();
		 driver.manage().window().maximize();
		 driver.get("https://opensource-demo.orangehrmlive.com/index.php/auth/validateCredentials");
		 loginPage = new LoginPage();
  }
  
  @AfterMethod
  public void tearDown() {
	  driver.close();
  }           //public static void main(String[] args) {
		//BaseClass base = new BaseClass();
		//base.setUp();
		//base.setDown();
}



