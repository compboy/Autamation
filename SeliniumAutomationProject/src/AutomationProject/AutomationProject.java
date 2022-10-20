package AutomationProject;


	
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeUnit;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DateFormatSymbols;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class AutomationProject {

	public static void main(String[] args) throws Exception {
		
		
		
		

		File src =new File("D:\\BRACU\\PROGRAMMING\\Silenium\\ExcelFiles\\Excel.xlsx");
		FileInputStream fis = new FileInputStream(src);
		XSSFWorkbook xsf = new XSSFWorkbook(fis);
		
		//set the path of our executable browser drive.
		System.setProperty("webdriver.chrome.driver", "D:\\BRACU\\PROGRAMMING\\Silenium\\Driver & Library//chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		
		//Getting day 0f the week {0>Sunday,1>Monday,2>Tuesday,3>Wednesday,4>Thursday,5>Friday,6>Saturday}
		String [] WeekDayArray= {"Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"}; 
		        
		Date d=new Date();
		int day = d.getDay();
		System.out.println("Today  is "+WeekDayArray[day]+".");
		
		//Geetting the sheet number as per week day.
		XSSFSheet WeekDay = xsf.getSheetAt(day);
		int rowCount =WeekDay.getLastRowNum();
		
		System.out.println("Total row is "+rowCount);
		List<String> TempKeywords = new ArrayList<String>();
		
		for(int i=2;i<=rowCount;i++) {
			
			String keyword = WeekDay.getRow(i).getCell(2).getStringCellValue();
			System.out.println(i+"th Row keyword is "+ keyword);
			TempKeywords.add(keyword);
		}
		  System.out.println();
		xsf.close();

		
		
		//driver.manage().window().maximize();
		for(int i=0;i<TempKeywords.size();i++) {
		System.out.print("Finding the WebElements for "+TempKeywords.get(i));
		driver.get("https://www.google.com");
		Thread.sleep(2000);
		driver.findElement(By.name("q")).sendKeys(TempKeywords.get(i));
		Thread.sleep(2000);
		
		WebElement list;
		List <WebElement > list1 =driver.findElements(By.xpath("//ul[@role='listbox']/li/descendant::div[@class='wM6W7d']"));
        List<String> Keywords = new ArrayList<String>();
		
		//Getting the search list
		for(WebElement String: list1) {
			String StringName = String.getText().trim();
			//System.out.println(StringName);
			Keywords.add(StringName);
		}
		
		
		  //Finding the longest Keywords 
		 String longestString = Keywords .stream().max(Comparator.comparingInt(String::length)).get();
		  System.out.println("\nLongest String is = " + longestString);
		  
		  //Smallest Keywords 
		  String smallestString = Keywords .stream().min(Comparator.comparingInt(String::length)).get();
		  System.out.println("Smallest String is = " + smallestString);
		  
		  writeData(longestString,smallestString, i, day);
		  Thread.sleep(1000);
		  System.out.println();
		 
		}
	}

	private static void writeData(String S1, String S2, int index, int day) throws Exception {

		File src =new File("D:\\BRACU\\PROGRAMMING\\Silenium\\ExcelFiles\\Excel.xlsx");
		FileInputStream fis = new FileInputStream(src);
		
		XSSFWorkbook xsf = new XSSFWorkbook(fis);
		
		XSSFSheet Week_day = xsf.getSheetAt(day);
		
		Week_day.getRow(index+2).createCell(3).setCellValue(S1);
		Week_day.getRow(index+2).createCell(4).setCellValue(S2);
		FileOutputStream fout = new FileOutputStream(src);
		xsf.write(fout);
		xsf.close();
		
	}

}
