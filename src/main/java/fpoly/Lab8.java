package fpoly;

import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;

import javax.swing.text.Element;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.By.ByName;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class Lab8 {
WebDriver driver;


public UIMap uimap;
public UIMap datafile;
//ghi file xlxs
Workbook workbook;
Sheet sheet; 
Map<String, Object[]> TestNGResult;
public static String driverPath = "C:\\jav\\selenium-java-3.141.59";
 @Test
 public void test1() {
	 workbook = new XSSFWorkbook(); // .xlsx
	 sheet = workbook.createSheet("KetQua");
	 TestNGResult = new LinkedHashMap<String, Object[]>();
	 System.setProperty("webdriver.chrome.driver", "D:\\chromedriver.exe");
	 driver = new ChromeDriver();
	 
 }	
	
@Test(priority = 1)
public void test2() {
	uimap = new UIMap("C:\\Users\\trung\\eclipse-workspace\\trung_poly_selenium\\Data\\name.properties");
	datafile= new UIMap("C:\\Users\\trung\\eclipse-workspace\\trung_poly_selenium\\Data\\dulieu.properties");
try {
	
	String url = "http://localhost:8080/Trung_Assignment_KTNC/login.htm";

	driver.get(url);
	
//
	WebElement username = driver.findElement(uimap.getLocator("user"));
	WebElement pass = driver.findElement(uimap.getLocator("pass"));
	
	
	//su dung DU LIEU
	username.click();
	username.sendKeys(datafile.getData("username"));
	pass.click();
	pass.sendKeys(datafile.getData("password"));
	driver.findElement(uimap.getLocator("login")).click();;
	
	//dat ten cho excel
	TestNGResult.put("0", new Object[] {
			"Ket qua mong muon","Thuc te","stauts"
		});
	if(!driver.getCurrentUrl().equals(url)) {
		TestNGResult.put("1", new Object[] {
				"Dang nhap thang cong","pass","Passed"
			});
	}else {
		TestNGResult.put("1", new Object[] {
				"Dang nhap thang cong","fail","Failed"
			});
	}
} catch (Exception e) {
	// TODO: handle exception
}
}






@Test(priority = 2)
public void writer() {
	// Excel username,password,kq
	Set<String> keySet = TestNGResult.keySet();
	int rown = 1;
	for(String key : keySet) {
		Row row = sheet.createRow(rown++);
		Object[] objarr = TestNGResult.get(key);
		int sheetn = 1;
		//set do rong cua file
		sheet.setColumnWidth(1, 7000);
		sheet.setColumnWidth(2, 7000);
		sheet.setColumnWidth(3, 7000);
		for(Object obj : objarr) {
			Cell cell = row.createCell(sheetn++);
			if(obj instanceof Date) {
				cell.setCellValue((Date) obj);
			}else 
			if(obj instanceof Boolean) {
				cell.setCellValue((Boolean) obj);
			}
			else 
				if(obj instanceof String) {
					cell.setCellValue((String) obj);
				}
				else 
					if(obj instanceof Double) {
						cell.setCellValue((Double) obj);
					}
		}
	}
	try {
		
		FileOutputStream out = new FileOutputStream(new File("Test.xlsx"));
		workbook.write(out);
		Desktop.getDesktop().open(new File("Test.xlsx"));
		out.close();
		System.out.println("Ghi thanh cong");
		driver.close();
	} catch (Exception e) {
		// TODO: handle exception
		System.out.println("That bai");
	}
	

}


}
