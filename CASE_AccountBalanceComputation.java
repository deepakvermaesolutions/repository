

import java.time.Duration;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;



public class CASE_AccountBalanceComputation
{

	WebDriver driver;
	JavascriptExecutor js;
	 WebDriverWait wait;
	 String caseIDCreated;
	 
	
	@Test(priority=1)
	public void individualCustomercreditCardCaseJourney() throws InterruptedException
	{
		String baseUrl="https://sbcdev.crmnext.com/accg7/app/login/login";
//		String driverPath="D:\\Ashish\\Projects\\Automation\\Drivers\\chromedriver_win32\\chromedriver.exe";
//		String baseUrl="https://sbcdev.crmnext.com/sag7/app/login/login";
		String driverPath="D:\\Deepak\\Drivers\\ChromeDriver\\chromedriver.exe";
		String BBN="10000000537";
		
		System.setProperty("webdriver.chrome.driver",driverPath);
   // 	WebDriver driver = new ChromeDriver();
    	driver = new ChromeDriver();
    //	driver.manage().deleteAllCookies();
    //	JavascriptExecutor js = (JavascriptExecutor) driver;
        js = (JavascriptExecutor) driver;
    	driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
   //     String baseUrl = "https://sbcdev.crmnext.com/sn/app/login/login";
   //     WebDriverWait wait= new WebDriverWait(driver, Duration.ofSeconds(10));
        wait= new WebDriverWait(driver, Duration.ofSeconds(10));
        
        driver.get(baseUrl);    
        
    	//driver.manage().deleteAllCookies();
    
    
   //     String baseUrl = "https://sbcdev.crmnext.com/sn/app/login/login";
     
      WebElement username=driver.findElement(By.id("TxtName"));
      WebElement password=driver.findElement(By.id("TxtPassword"));
      
      username.sendKeys("100082");
      password.sendKeys("acid_qa");
      driver.findElement(By.name("command")).click();
//      Thread.sleep(3000);
      //Click Quick Links and Go to Customer Search
      
      Actions quick = new Actions(driver);
      quick.moveToElement(driver.findElement(By.className("icon-link"))).perform();
//      Thread.sleep(3000);
      
      driver.findElement(By.xpath("//*[@id=\"QuickLinksDiv\"]/ul/li[1]/a/span[1]")).click();
      System.out.println("Click Quick Links");
//      Thread.sleep(3000);
      
      //Customer Search Landing Page

      System.out.println("Customer Search Landing Page");
    
      WebElement inputbbn=driver.findElement(By.name("100000012"));
      inputbbn.sendKeys(BBN); //25600000469 
      System.out.println("BBN Entered");
      driver.findElement(By.xpath("//*[@id=\"object-action-button\"]/div/a[1]")).click();
      Thread.sleep(2000);
      wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//label[@for='checkbox0']")));
      driver.findElement(By.xpath("//label[@for='checkbox0']")).click();
      Thread.sleep(3000);
      
      
      driver.findElement(By.xpath("//a[@title='Next']")).click();
    
//      Thread.sleep(2000);
      js.executeScript("window.scrollBy(0,1000)","" );
      
      wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@data-autoid='WidgetActionButton']")));
      driver.findElement(By.xpath("//*[@data-autoid='WidgetActionButton']")).click(); //gold 7
      System.out.println("Click on Fetch Button");
      
      System.out.println("Customer 360 Landing Page");
      wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(@title,'click here to see more actions')]")));
      driver.findElement(By.xpath("//*[contains(@title,'click here to see more actions')]")).click();
    //  driver.findElement(By.xpath("//div[contains(@title,'click here to see more actions')]")).click();
      System.out.println("Clicked on Custom Action Button in Customer 360");
      
//      Thread.sleep(2000);
      List<WebElement> cust350CustumButton=driver.findElements(By.cssSelector(".mb-24.flex.items-center.acd-link"));
      for(int i=0;i<cust350CustumButton.size();i++)
      {
      	String customActionButtonList =cust350CustumButton.get(i).getText();
      	if(customActionButtonList.equalsIgnoreCase("start your journey"))
      	{
      		//driver.findElements(By.cssSelector(".mb-24.flex.items-center.acd-link")).get(i).click();
      		cust350CustumButton.get(i).click();
      	}
      }
      
      
      System.out.println("Clicked on Start your journey button in Customer 360 for Case creation");
      
//      Thread.sleep(2000);
      
    //Get handles of the windows
      String mainWindowHandle = driver.getWindowHandle();

//      Thread.sleep(3000);
      Set<String> allWindowHandles = driver.getWindowHandles();
      Iterator<String> iterator=	allWindowHandles.iterator();
//     String parentWindow= iterator.next();
//     String childWindow= iterator.next();
      System.out.println("Handle Child Windows");
   
      while (iterator.hasNext()) 
      {
          String ChildWindow = iterator.next();
              if (!mainWindowHandle.equalsIgnoreCase(ChildWindow)) 
              {
              driver.manage().window().maximize();
              driver.switchTo().window(ChildWindow);
//              Thread.sleep(2000);
              WebElement CustomerService = driver.findElement(By.xpath("//*[@id=\"2\"]"));
   
              CustomerService.click();	
              }    
      }

//   driver.switchTo().window(mainWindowHandle);
   wait.until(ExpectedConditions.presenceOfElementLocated(By.name("CASE_SUBCATEGORY1")));
     WebElement SUBCategroy=driver.findElement(By.name("CASE_SUBCATEGORY1"));
     SUBCategroy.sendKeys("Account Balance Computation");
     Thread.sleep(1000);
     SUBCategroy.sendKeys(Keys.ARROW_DOWN);
     SUBCategroy.sendKeys(Keys.ENTER);
     Thread.sleep(1000);
     
     wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[@data-autoid='CASE_PRDHOLDINGNUMBER_srch']")));
     driver.findElement(By.xpath("//a[@data-autoid='CASE_PRDHOLDINGNUMBER_srch']")).click();
     
     List<WebElement> creditcardproduct=driver.findElements(By.xpath("//div[contains(@data-autoid,'Name')]"));
     
     for(int i=0;i<creditcardproduct.size();i++)
     {
    	 String creditCardProductName=creditcardproduct.get(i).getText();
    	 if(creditCardProductName.equalsIgnoreCase("MasterCard CashBack Card"))
    	 {
    		 creditcardproduct.get(i).click();
    		 System.out.println("Credit Card Account Number is selected");
    		 break;
    		 
    	 }
    	 else
    	 {
    		 System.out.println("Credit Card Product not found");
    	 }
     }
   
     
     Thread.sleep(3000);
     Select  SourceValues= new Select (driver.findElement(By.xpath("//select[@name='CASE_ORIGIN']")));
     System.out.println(SourceValues);
     
     SourceValues.selectByVisibleText("Call Center");
     System.out.println("Call Center is selected");
     Thread.sleep(3000);
     
     try 
     {
     driver.findElement(By.xpath("//a[@data-autoid='FlowNext']")).click();
     }
     catch(StaleElementReferenceException e)
     {
    	 wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//a[@data-autoid='FlowNext']")));
    	 driver.findElement(By.xpath("//a[@data-autoid='FlowNext']")).click();
     }
     
   //  driver.navigate().refresh();
     js.executeScript("window.scrollBy(0,1500)", "");
    driver.findElement(By.xpath("//*[@data-autoid='cust_878_ctrl']")).sendKeys("Test Details of Concern Field");
    System.out.println("Entered text in Details of Concern field");
    
    
    WebElement computationRequestDropdownElement = driver.findElement(By.xpath("//*[@data-autoid='cust_1896_ctrl']"));
    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//select[@data-autoid='cust_1896_ctrl']")));
    Thread.sleep(5000);
    Select  computationRequestDropdown= new Select(computationRequestDropdownElement);
    
    computationRequestDropdown.selectByValue("Available Balance Computation");
    
    
    System.out.println("Selected Available Balance Computation from dropdown");
    
    driver.findElement(By.xpath("//*[@data-autoid='cust_471_ctrl']")).sendKeys("Test Remarks");
    System.out.println("Remarks Entered");
    
    driver.findElement(By.xpath("//a[@data-autoid='FlowNext']")).click();
    System.out.println("Clicked on Save and Proceed button");
    
    WebElement caseID=driver.findElement(By.xpath("//div[@class='summaryband__item flex items-center']/div/div/div/span[@data-autoid='CASE_NUMBER_ctrl']"));
    caseIDCreated=caseID.getText();  
    System.out.println("Account Balance Computation CaseID:"+caseIDCreated);
    
    
    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[contains(@class,'header__item header__profile')]/a/img")));
    WebElement logoutElement=driver.findElement(By.xpath("//div[contains(@class,'header__item header__profile')]/a/img"));
    
    JavascriptExecutor jse= (JavascriptExecutor)driver;
    jse.executeScript("arguments[0].click()", logoutElement);

    Thread.sleep(2000);
    System.out.println("Clicked on right Top cornor for logout");
    
    driver.findElement(By.xpath("//a[@title='Logout']")).click();
    Thread.sleep(2000);
    System.out.println("Logout Successfully");
    
    
    
	}
	
	@Test(priority=2)
	public void stage1() throws InterruptedException 
	{

		driver.findElement(By.name("UserName")).sendKeys("203953");
    	Thread.sleep(2000);
        driver.findElement(By.name("SecureTextBox.Text")).sendKeys("acid_qa");
   
        driver.findElement(By.name("command")).click();
   
        System.out.println("CAC Authorizer Login Successfull");
        Thread.sleep(4000);
        driver.findElement(By.className("icon-service")).click();
        System.out.println("Click Service");
   
        Thread.sleep(3000);
        driver.findElement(By.className("icon-obj9")).click();
        System.out.println("Click Cases");
   
        System.out.println("Cases Home Page Open");
        
        Thread.sleep(3000);
     // This  will scroll down the page by  1000 pixel vertical	
        try
        {        
        	Thread.sleep(2000);
        	js.executeScript("window.scrollBy(0,500)");
        
        
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("QueryViewId")));
   
    Select View = new Select(driver.findElement(By.name("QueryViewId")));
    
        View.selectByVisibleText("Assigned to Department");
        Thread.sleep(3000);
        driver.findElement(By.xpath("//a[@class='filterGroup__button acid-btn acid-btn--outline-brand']")).click();
        }
        catch(StaleElementReferenceException e)
        {
        	Thread.sleep(3000);
        	js.executeScript("window.scrollBy(0,500)");
            
            
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("QueryViewId")));
       
        Select View = new Select(driver.findElement(By.name("QueryViewId")));
        
            View.selectByVisibleText("Assigned to Department");

            driver.findElement(By.xpath("//a[@class='filterGroup__button acid-btn acid-btn--outline-brand']")).click();
            Thread.sleep(2000);
        }

//     String CaseIDStage1=caseIDcasting;
       String CaseActual=caseIDCreated.replaceFirst("^0+(?!$)", "");
       List<WebElement> caseSelectiNmyBucket=driver.findElements(By.xpath("//div[contains(@data-autoid,'CAS_EX1_147')]"));
      
        for(int i=0;i<caseSelectiNmyBucket.size();i++)
        {
        	String caseID2=caseSelectiNmyBucket.get(i).getText();
        	if(caseID2.equals(CaseActual))
        	{
     		
        		//wait.until(ExpectedConditions.stalenessOf(driver.findElement(By.cssSelector(".react-grid-Cell.react-grid-Cell--frozen.rdg-last--frozen"))));
				driver.findElement(By.cssSelector(".react-grid-Cell.react-grid-Cell--frozen.rdg-last--frozen")).click();
				System.out.println("Checkbox against cases is selected");
        		break;
        	}
        }
        
        
        driver.findElement(By.xpath("//*[contains(@title,'click here to see more actions')]")).click();
      System.out.println("Clicked on Custom Action Button in Customer 360");

   List<WebElement> cust350CustumButtonView=driver.findElements(By.cssSelector(".mb-24.flex.items-center.acd-link"));
      for(int i=0;i<cust350CustumButtonView.size();i++)
      {
      	String customActionButtonList =cust350CustumButtonView.get(i).getText();
      	if(customActionButtonList.equalsIgnoreCase("Self assign cases"))
      	{
      		//driver.findElements(By.cssSelector(".mb-24.flex.items-center.acd-link")).get(i).click();
      		cust350CustumButtonView.get(i).click();
      		System.out.println("Self Assigned Successfully");
      		break;
      		 
      	}
      }
driver.navigate().refresh();
try {

Select view1=new Select(driver.findElement(By.xpath("//select[@data-autoid='QueryViewId_ctrl']")));
view1.selectByValue("56");
System.out.println("Assigned to Me view selected");
}
catch(StaleElementReferenceException e)
{
	
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//select[@data-autoid='QueryViewId_ctrl']")));
		Select view1=new Select(driver.findElement(By.xpath("//select[@data-autoid='QueryViewId_ctrl']")));
		view1.selectByValue("56");
		System.out.println("Assigned to Me view selected");
	}

        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@class='react-grid-Row react-grid-Row--even']/div[@class='react-grid-Cell']")));
   
      List<WebElement> caseAssignedToMe= driver.findElements(By.xpath("//div[@class='react-grid-Row react-grid-Row--even']/div[@class='react-grid-Cell']"));
     
      for(int i=0;i<caseAssignedToMe.size();i++)
      {
    	  String editCaseID=caseAssignedToMe.get(i).getText();
    	  if(editCaseID.equals(caseIDCreated))
    		  
    	  {
    		  try {
    		  driver.findElement(By.xpath("//div[@class='react-grid-Row react-grid-Row--even']/div[@class='react-grid-Cell']")).click();
    		  System.out.println("Click Case ID");
    		  break;
    		  }
    		  catch(StaleElementReferenceException e)
    		  {
    			  driver.findElement(By.xpath("//div[@class='react-grid-Row react-grid-Row--even']/div[@class='react-grid-Cell']")).click();
        		  System.out.println("Click Case ID");
        		  break;
    		  }
    		  }
      }
      
      
      
      driver.findElement(By.xpath("//div[contains(@title,'click here to see more actions')]")).click();
  System.out.println("Clicked on Custom Action Button in Customer 360");
//Code of clicking on Edit button against case
 List<WebElement> cust350CustumButtononCase=driver.findElements(By.cssSelector(".mb-24.flex.items-center.acd-link"));
      wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector(".mb-24.flex.items-center.acd-link")));
    for(int i=0;i<cust350CustumButtononCase.size();i++)
    {
    	String customActionButtonList =cust350CustumButtononCase.get(i).getText();
    	if(customActionButtonList.equalsIgnoreCase("Edit"))
    	{
    		//driver.findElements(By.cssSelector(".mb-24.flex.items-center.acd-link")).get(i).click();
    		cust350CustumButtononCase.get(i).click();
    		System.out.println("Case Opened in Edit Mode");
    		break;
    		 
    	}
    }

        
    Select Department = new Select(driver.findElement(By.xpath("//select[@name='cust_512']")));
    Thread.sleep(2000);
    Department.selectByVisibleText("Accept");
    System.out.println("Seelct Department Decision");
    Thread.sleep(3000);
    
    driver.findElement(By.name("cust_471")).sendKeys("Remarks by Processor");
    Thread.sleep(3000);
    
    driver.findElement(By.xpath("//*[@id=\"object-action-button\"]/div/a[2]")).click();
    System.out.println("Click Save and Proceed Button");
    
    String CaseStatus=driver.findElement(By.xpath("//span[@data-autoid='CASE_STATUSCODE_ctrl']")).getText();
    if(CaseStatus.equalsIgnoreCase("Closed"))
    {
    System.out.println("Case Closed Successfully and its status is: "+CaseStatus);
    }	
    else 
    {
    	System.out.println("Case Status is not closed and It's current Status is: "+CaseStatus);
    }
    
    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[contains(@class,'header__item header__profile')]/a/img")));
    WebElement logoutElement=driver.findElement(By.xpath("//div[contains(@class,'header__item header__profile')]/a/img"));
    
    JavascriptExecutor jse= (JavascriptExecutor)driver;
    jse.executeScript("arguments[0].click()", logoutElement);

    Thread.sleep(2000);
    System.out.println("Clicked on right Top cornor for logout");
    
    driver.findElement(By.xpath("//a[@title='Logout']")).click();
    Thread.sleep(2000);
    System.out.println("Logout Successfully");
	}
	
	@Test(priority=3)
	public void stage2() throws InterruptedException 
	{

		driver.findElement(By.name("UserName")).sendKeys("150153");
    	Thread.sleep(2000);
        driver.findElement(By.name("SecureTextBox.Text")).sendKeys("acid_qa");
   
        driver.findElement(By.name("command")).click();
   
        System.out.println("CCG Authorizer Login Successfull");
        Thread.sleep(4000);
        driver.findElement(By.className("icon-service")).click();
        System.out.println("Click Service");
   
        Thread.sleep(3000);
        driver.findElement(By.className("icon-obj9")).click();
        System.out.println("Click Cases");
   
        System.out.println("Cases Home Page Open");
        
     // This  will scroll down the page by  1000 pixel vertical	
        try
        {        
        	js.executeScript("window.scrollBy(0,500)");
        
        
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("QueryViewId")));
   
    Select View = new Select(driver.findElement(By.name("QueryViewId")));
    
        View.selectByVisibleText("Assigned to Department");

        driver.findElement(By.xpath("//a[@class='filterGroup__button acid-btn acid-btn--outline-brand']")).click();
        }
        catch(StaleElementReferenceException e)
        {
        	
        	js.executeScript("window.scrollBy(0,500)");
            
            
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("QueryViewId")));
       
        Select View = new Select(driver.findElement(By.name("QueryViewId")));
        
            View.selectByVisibleText("Assigned to Department");

            Thread.sleep(3000);
            driver.findElement(By.xpath("//a[@class='filterGroup__button acid-btn acid-btn--outline-brand']")).click();
        	
        }

//     String CaseIDStage1=caseIDcasting;
       String CaseActual=caseIDCreated.replaceFirst("^0+(?!$)", "");
       List<WebElement> caseSelectiNmyBucket=driver.findElements(By.xpath("//div[contains(@data-autoid,'CAS_EX1_147')]"));
      
        for(int i=0;i<caseSelectiNmyBucket.size();i++)
        {
        	String caseID2=caseSelectiNmyBucket.get(i).getText();
        	if(caseID2.equals(CaseActual))
        	{
     		
        		Thread.sleep(2000);
        		//wait.until(ExpectedConditions.stalenessOf(driver.findElement(By.cssSelector(".react-grid-Cell.react-grid-Cell--frozen.rdg-last--frozen"))));
				driver.findElement(By.cssSelector(".react-grid-Cell.react-grid-Cell--frozen.rdg-last--frozen")).click();
				System.out.println("Checkbox against cases is selected");
        		break;
        	}
        }
        
        
        driver.findElement(By.xpath("//*[contains(@title,'click here to see more actions')]")).click();
      System.out.println("Clicked on Custom Action Button in Customer 360");

   List<WebElement> cust350CustumButtonView=driver.findElements(By.cssSelector(".mb-24.flex.items-center.acd-link"));
      for(int i=0;i<cust350CustumButtonView.size();i++)
      {
      	String customActionButtonList =cust350CustumButtonView.get(i).getText();
      	if(customActionButtonList.equalsIgnoreCase("Self assign cases"))
      	{
      		//driver.findElements(By.cssSelector(".mb-24.flex.items-center.acd-link")).get(i).click();
      		cust350CustumButtonView.get(i).click();
      		System.out.println("Self Assigned Successfully");
      		break;
      		 
      	}
      }
driver.navigate().refresh();
try {

Select view1=new Select(driver.findElement(By.xpath("//select[@data-autoid='QueryViewId_ctrl']")));
view1.selectByValue("56");
System.out.println("Assigned to Me view selected");
}
catch(StaleElementReferenceException e)
{
	
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//select[@data-autoid='QueryViewId_ctrl']")));
		Select view1=new Select(driver.findElement(By.xpath("//select[@data-autoid='QueryViewId_ctrl']")));
		view1.selectByValue("56");
		System.out.println("Assigned to Me view selected");
	}

        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@class='react-grid-Row react-grid-Row--even']/div[@class='react-grid-Cell']")));
   
      List<WebElement> caseAssignedToMe= driver.findElements(By.xpath("//div[@class='react-grid-Row react-grid-Row--even']/div[@class='react-grid-Cell']"));
     
      for(int i=0;i<caseAssignedToMe.size();i++)
      {
    	  String editCaseID=caseAssignedToMe.get(i).getText();
    	  if(editCaseID.equals(caseIDCreated))
    		  
    	  {
    		  try {
    		  driver.findElement(By.xpath("//div[@class='react-grid-Row react-grid-Row--even']/div[@class='react-grid-Cell']")).click();
    		  System.out.println("Click Case ID");
    		  break;
    		  }
    		  catch(StaleElementReferenceException e)
    		  {
    			  driver.findElement(By.xpath("//div[@class='react-grid-Row react-grid-Row--even']/div[@class='react-grid-Cell']")).click();
        		  System.out.println("Click Case ID");
        		  break;
    		  }
    		  }
      }
      
      
      
      driver.findElement(By.xpath("//div[contains(@title,'click here to see more actions')]")).click();
  System.out.println("Clicked on Custom Action Button in Customer 360");
//Code of clicking on Edit button against case
 List<WebElement> cust350CustumButtononCase=driver.findElements(By.cssSelector(".mb-24.flex.items-center.acd-link"));
      wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector(".mb-24.flex.items-center.acd-link")));
    for(int i=0;i<cust350CustumButtononCase.size();i++)
    {
    	String customActionButtonList =cust350CustumButtononCase.get(i).getText();
    	if(customActionButtonList.equalsIgnoreCase("Edit"))
    	{
    		//driver.findElements(By.cssSelector(".mb-24.flex.items-center.acd-link")).get(i).click();
    		cust350CustumButtononCase.get(i).click();
    		System.out.println("Case Opened in Edit Mode");
    		break;
    		 
    	}
    }

        
    Select Department = new Select(driver.findElement(By.xpath("//select[@name='cust_512']")));
    Thread.sleep(2000);
    Department.selectByVisibleText("Accept");
    System.out.println("Seelct Department Decision");
    Thread.sleep(3000);
    
    driver.findElement(By.name("cust_471")).sendKeys("Remarks by Processor");
    Thread.sleep(3000);
    
    driver.findElement(By.xpath("//*[@id=\"object-action-button\"]/div/a[2]")).click();
    System.out.println("Click Save and Proceed Button");
    
    String CaseStatus=driver.findElement(By.xpath("//span[@data-autoid='CASE_STATUSCODE_ctrl']")).getText();
    if(CaseStatus.equalsIgnoreCase("Closed"))
    {
    System.out.println("Case Closed Successfully and its status is: "+CaseStatus);
    }	
    else 
    {
    	System.out.println("Case Status is not closed and It's current Status is: "+CaseStatus);
    }
	}
	
//	@AfterTest
//	public void quitBroweser() 
//	{
//		System.out.println("Case Journey Completed Successfully");
//		driver.quit();
//		
//	}
	
}
