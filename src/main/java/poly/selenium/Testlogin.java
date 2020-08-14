package poly.selenium;
import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By.ByName;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.Assert;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import poly.enity.LIST;
import poly.enity.Staffs;
import poly.enity.depart;
import poly.enity.users;


public class Testlogin {
	public WebDriver driver;
	public String working;
	Workbook workbook;
	Sheet sheer;
	Map<String, Object[]> TestNGResults;
	LIST li = new LIST();
	public static String driverPath = "C:\\Users\\Admin\\Desktop\\activation\\jav[ZaloPC_Folder]\\selenium-java-3.141.59";
	// dung de toal tong hop luot pass va fail cho moi lan test 
	List<String> demPass = new ArrayList<String>();
	List<String> demFail = new ArrayList<String>();

	// thiết lập cho file exel chạy đầu tiên
	@BeforeClass(alwaysRun = true)
	public void suitsetup() {
		workbook = new XSSFWorkbook();
		sheer = workbook.createSheet("Ket Qua Kiểm Tra");
		TestNGResults = new LinkedHashMap<String, Object[]>();

		System.setProperty("webdriver.chrome.driver", "D:\\chromedriver.exe");
		
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(6, TimeUnit.SECONDS);
	}
	// test case đăng nhập
		@Test(groups = "selenium", priority = 1)
		public void Test() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
			try {
				// link truy cap vao trang web để test
				String url = "http://localhost:8008/Trung_Assignment_KTNC/login.htm";
				driver.manage().window().maximize();
				//tieu de cua tung cot trong exel
				TestNGResults.put("1", new Object[] { "ID", "Test Data", "Expected Results", "Actual Results", "Status"

				});

				List<users> list = li.list(0);

				for (int i = 0; i < list.size(); i++) {
					// lay dữ lieu duong truyen t rang web
					driver.get(url);
					//lay du lieu tu exel
					driver.findElement(ByName.name("username")).click();
					;
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					driver.findElement(ByName.name("username")).sendKeys(list.get(i).getUsername());
					driver.findElement(ByName.name("pass")).click();
					;
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					driver.findElement(ByName.name("pass")).sendKeys(list.get(i).getPassword());
					;
					;
					driver.findElement(ByName.name("login")).click();
					;
					// Ket mong muon thong

					// ket qua mong muon dang nhap thanh cong
					//lay du lieu xong do len bang exel
					if (!driver.getCurrentUrl().equals(url)) {
						TestNGResults.put("2" + i,
								new Object[] { "Test_Login_" + (i + 1),
										"Login \t\n with Username = '" + list.get(i).getUsername() + "'\t\n Password = '"
												+ list.get(i).getPassword() + "'",
										"Đăng Nhập thành công", "PASS", "PASSED" });
						demPass.add("PASSED");
					} else {
						TestNGResults.put("3" + i,
								new Object[] { "Test_Login_" + (i + 1),
										"Login \t\n with Username = '" + list.get(i).getUsername() + "'\t\n Password = '"
												+ list.get(i).getPassword() + "'",
										"Đăng Nhập thành công",
										driver.findElement(ByName.xpath("/html/body/div/div[2]/div[1]/div[1]")).getText(),
										"FAILED" });
						demFail.add("Faild");				}
					Assert.assertTrue(true);
				}

			} catch (Exception e) {
				// TODO: handle exception
				TestNGResults.put("4", new Object[] { "Error", "Fill login from data (Username/Password)",
						"Login details gets filled", "Fail" });
				Assert.assertTrue(false);
				System.out.println(e);
			} finally {
				/// html/body/div/div[1]/div[2]/ul/li[2]/a

			}

		}
		// test case Login
		@Test(groups = "selenium", priority = 6)
		public void Test5() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
			try {
				driver.get("http://localhost:8008/Trung_Assignment_KTNC/depart/");
				List<depart> list_depart = li.list_depart(3);
				int passed = 0;
				for (int i = 0; i < list_depart.size(); i++) {
					driver.get("http://localhost:8008/Trung_Assignment_KTNC/depart/");
					driver.findElement(ByName.name("id")).click();
					;
					driver.findElement(ByName.name("id")).sendKeys(list_depart.get(i).getId());
					driver.findElement(ByName.name("name")).click();
					;
					driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
					driver.findElement(ByName.name("name")).sendKeys(list_depart.get(i).getName());
					;
					driver.findElement(ByName.name("adddepart")).click();
					driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);

					Thread.sleep(1000);
					String error = driver.findElement(ByName.xpath("/html/body/div[2]/span[3]")).getText();
					System.out.println(error);

					List<users> count = li.list(0);
					List<users> count2 = li.list(1);
					int soTest = (count.size() + count2.size() * 2 + count.size() + list_depart.size());
					if (error.equals("Vui lòng kiểm tra các lỗi nhập liệu")) {
						TestNGResults.put("2" + (soTest + i),
								new Object[] { "Test_QLPB_" + (i + 1),
										"Add with ID : " + list_depart.get(i).getId() + " and Depart name: "
												+ list_depart.get(i).getName(),
										"Add Phòng Ban thành công", error, "FAILED" });
						demFail.add("Faild");
					} else {
						TestNGResults.put("2" + (soTest + i),
								new Object[] { "Test_QLPB_" + (i + 1),
										"Add with ID : " + list_depart.get(i).getId() + " and Depart name: "
												+ list_depart.get(i).getName(),
										"Add Phòng Ban thành công", error, "PASSED" });
						demPass.add("PASSED");
					}

				}

			} catch (Exception e) {
				// TODO: handle exception
				TestNGResults.put("4", new Object[] { "Error", "Fill login from data (Username/Password)",
						"Login details gets filled", "Fail" });
				Assert.assertTrue(false);
				System.out.println(e);
			}

		}




}
