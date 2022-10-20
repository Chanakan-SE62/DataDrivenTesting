import static org.junit.jupiter.api.Assertions.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

class UC1CheckLogin {

	@Test
	void testCheckLogin() throws Exception {
		System.setProperty("webdriver.chrome.driver", "C:/chromedriver.exe");

	    SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");  
	    Date date = new Date();  
		String testDate = formatter.format(date);
		String testerName = "Chanakan Chaisena";
	   
		String path = "C:\\Users\\chana\\Desktop\\testdata.xlsx";
		FileInputStream fs = new FileInputStream(path);

		//Creating a workbook
		XSSFWorkbook workbook = new XSSFWorkbook(fs);
		XSSFSheet sheet = workbook.getSheetAt(0);
		int row = sheet.getLastRowNum()+1;
		//System.out.println(row);
		

		for(int i = 1; i<row; i++) {
			WebDriver driver = new ChromeDriver();
			driver.get("https://katalon-demo-cura.herokuapp.com/");
			driver.findElement(By.id("btn-make-appointment")).click();
			String Test_Case_ID = sheet.getRow(i).getCell(0).toString();
			String Username;
			String Password;
			if(Test_Case_ID.equals("tc104")) {
				Username = "";
				Password = "";
			}else {
				Username = sheet.getRow(i).getCell(1).toString();			
				Password = sheet.getRow(i).getCell(2).toString();		
			}
			driver.findElement(By.id("txt-username")).sendKeys(Username);
			driver.findElement(By.id("txt-password")).sendKeys(Password);
			driver.findElement(By.id("btn-login")).click();
			
			if(Test_Case_ID.equals("TC101")) {
				//System.out.println("tc101");
				String Actual = driver.findElement(By.xpath("/html/body/section/div/div/div/h2")).getText();
				String Expected = sheet.getRow(i).getCell(3).toString();
				Row rows = sheet.getRow(i);
				Cell cell = rows.createCell(4);
				cell.setCellValue(Actual);
				if(Actual.equals(Expected)) {
					Cell cell2 = rows.createCell(5);
					cell2.setCellValue("Pass");
				}else {
					Cell cell2 = rows.createCell(5);
					cell2.setCellValue("Fail");			
				}
				Cell cell3 = rows.createCell(6);
				cell3.setCellValue(testDate);
				Cell cell4 = rows.createCell(7);
				cell4.setCellValue(testerName);				
				FileOutputStream fos = new FileOutputStream(path);
				workbook.write(fos);
			} else {
				String Actual = driver.findElement(By.xpath("/html/body/section/div/div/div[1]/p[2]")).getText();
				String Expected = sheet.getRow(i).getCell(3).toString();
				Row rows = sheet.getRow(i);
				Cell cell = rows.createCell(4);
				cell.setCellValue(Actual);
				if(Actual.equals(Expected)) {
					Cell cell2 = rows.createCell(5);
					cell2.setCellValue("Pass");
				}else {
					Cell cell2 = rows.createCell(5);
					cell2.setCellValue("Fail");			
				}
				Cell cell3 = rows.createCell(6);
				cell3.setCellValue(testDate);
				Cell cell4 = rows.createCell(7);
				cell4.setCellValue(testerName);				
				FileOutputStream fos = new FileOutputStream(path);
				workbook.write(fos);
				
			}
			driver.close();
			System.out.println("ID = "+Test_Case_ID);
			System.out.println("ͼҹ = "+Username);
			System.out.println("ʼҹ = "+Password);
			
		}
		System.out.println("TestData Successfully");
	
	}

}