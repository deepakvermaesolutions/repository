package testpackage;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class ReadDataFromExcel {

	public static WebDriver driver;
	public static void main(String[] args) throws InterruptedException {
		
		String Username = null;
        String Password = null;
        // System Property for Chrome Driver   
     System.setProperty("webdriver.chrome.driver", "D:\\Deepak\\Drivers\\ChromeDriver\\chromedriver.exe");  
     
          // Instantiate a ChromeDriver class.     
      driver=new ChromeDriver();  
      
        // Launch Website  
    
     try {
         FileInputStream fStream = new FileInputStream(new File(
                 "D:\\Deepak\\Execl_File\\Username&Pwd.xlsx")); //Enter the path to your excel here

         // Create workbook instance referencing the file created above
         XSSFWorkbook workbook = new XSSFWorkbook(fStream);

         // Get sheet from the workbook
         XSSFSheet sheet = workbook.getSheetAt(0); // getting first sheet

         
         int rowCount = sheet.getLastRowNum();

         System.out.println("the no of rows are : " + rowCount);
         
         int lastCellNum = sheet.getRow(0).getLastCellNum();

         System.out.println("the no of cells are : " +lastCellNum);
         
         for (int i=1; i<=rowCount; i++)
         { 
        	
        		 Username = sheet.getRow(i).getCell(0).getStringCellValue();

                 Password = sheet.getRow(i).getCell(1).getStringCellValue();
                 
               // int order=(int) sheet.getRow(row).getCell(2).getNumericCellValue();

                System.out.println(Username + " , " + Password);
        	
                driver.get("https://ksctest.crmnext.com//sanew/app//login//login");  
                driver.manage().window().maximize(); // Maximize window
                driver.findElement(By.xpath("//*[@id=\"TxtName\"]")).sendKeys(Username);  //Enter user name
                driver.findElement(By.xpath("//*[@id=\"TxtPassword\"]")).sendKeys(Password); // Enter Password
                driver.findElement(By.xpath("//*[@id=\"registration\"]/input")).click();  //Click Login button
               
                String S = "Summary - CRMnext - Smart.Easy.Complete";
                XSSFCell cell = sheet.getRow(i).createCell(2);
                
                if(driver.getTitle().matches(S)) {
                	
                	System.out.println("Login successfully");
                	Thread.sleep(4000);
                    driver.findElement(By.xpath("//a[@class='header-profile__media min-wt-2 wt-2 ht-2 overflow-hidden']//img[@name='ProfileImage_header']")).click();
                    driver.findElement(By.xpath("//span[normalize-space()='Logout']")).click();
                // FileOutputStream outputStream = new FileOutputStream("D:\\\\Deepak\\\\Execl_File\\\\Username&Pwd.xlsx");
                //workbook.write(outputStream);
                cell.setCellValue("Pass");
                }
			
				 else { 
					 
					WebElement ele =  driver.findElement(By.xpath("//span[normalize-space()='Invalid User Name or Password.']"));
					System.out.println(ele.getText());
					 //System.out.println("Login Failed : Please input correct user name and password");
					 
					 cell.setCellValue("Fail"); }
				 
                FileOutputStream outputStream = new FileOutputStream("D:\\Deepak\\Execl_File\\Username&Pwd.xlsx");
                workbook.write(outputStream);
                
        
         fStream.close();
     } 
}
catch (Exception e) {
	
	e.printStackTrace();
         
     }
}
}
	
	


	
	
	
	