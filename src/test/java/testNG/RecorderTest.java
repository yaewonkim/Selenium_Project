package testNG;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.Assert;
import org.testng.ITestContext;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import PDFEmail.BaseClass;
import PDFEmail.JypersionListener;

@Listeners(JypersionListener.class)
public class RecorderTest extends BaseClass {
	
  public static String file_location = "C:\\Users\\YaewonKim\\Desktop\\TestNG.xls";
  static String SheetName= "Sheet2";	
  private StringBuffer verificationErrors = new StringBuffer();

  
  
  @BeforeClass(alwaysRun = true)
  public void setUp() throws Exception {
	System.setProperty("webdriver.gecko.driver", "C:\\Users\\YaewonKim\\workspace\\Selenium_Project\\drivers\\geckodriver.exe");
    driver = new FirefoxDriver();
    
//Chrome
//	System.setProperty("webdriver.chrome.driver", "C:\\Users\\YaewonKim\\workspace\\Selenium_Project\\drivers\\chromedriver.exe");
//	WebDriver driver = new ChromeDriver();
	
//Internet Explore
//	System.setProperty("webdriver.ie.driver", "C:\\Users\\YaewonKim\\workspace\\Selenium_Project\\drivers\\IEDriverServer.exe");
//	WebDriver driver = new InternetExplorerDriver();
    
    driver.manage().timeouts().implicitlyWait(12, TimeUnit.SECONDS);
    driver.manage().window().maximize();
  }
  
  
  @Test(dataProvider="SearchNameData")
  public void FirstTest(String Name) throws Exception {
	System.out.println("Name: "+Name);
    driver.get("http://examples.codecharge.com/EmployeeDirectory/Default.php");
    
    driver.findElement(By.name("filter")).click();
    driver.findElement(By.name("filter")).click();    
    driver.findElement(By.name("filter")).clear();
    driver.findElement(By.name("filter")).sendKeys(Name);
    driver.findElement(By.name("DoSearch")).click();
    driver.findElement(By.linkText(Name)).click();
    driver.findElement(By.xpath("(.//*[normalize-space(text()) and normalize-space(.)='Name'])[1]/following::td[1]")).click();
  }
  
  
  @Test
  public void SecondTest() throws Exception{
	  System.out.println("SecondTest입니다.");
  }
  
  

  @Test
  public void ThirdTest() throws Exception{
	  System.out.println("ThirdTest입니다.");
  }
  
  
  @AfterClass(alwaysRun = true)
  public void tearDownAfterClass() throws Exception {
    driver.quit();
    String verificationErrorString = verificationErrors.toString();
    if (!"".equals(verificationErrorString)) {
      Assert.fail(verificationErrorString);
    }
  }
  
  
  @AfterSuite
  public void tearDownAfterSuite(ITestContext context){
		 String suiteName = context.getName();
		 
		 Date today = new Date();
	        
	     SimpleDateFormat date = new SimpleDateFormat("yyyy/MM/dd");
	     SimpleDateFormat time = new SimpleDateFormat("HH:mm:ss");
	     
		 String title_time = "_"+date.format(today)+"_"+time.format(today);
		    
		sendPDFReportByGMail("sender's gmail address", "gmail password", "receiver's email address", "[TestAuto] 테스트 결과 안내"+title_time, "", suiteName);
  }
   
  
  @DataProvider(name="SearchNameData")
  public Object[][] SearchNameData() {
	  System.out.println("SearchNameData들어옴");
	  Object[][] arrayObject = getExcelData(file_location, SheetName);
	  return arrayObject;
  }

}
