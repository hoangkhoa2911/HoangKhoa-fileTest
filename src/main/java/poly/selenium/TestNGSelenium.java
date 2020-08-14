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

public class TestNGSelenium {

	public WebDriver driver;

	public String working;
	// link tham khao
	// https://viblo.asia/p/testng-data-provider-voi-excel-bJzKmMwPK9N?fbclid=IwAR2jp060WzlZcNA2rv8Hs6CmrRPQwcz8lqQJUtaYBkrlOLFSHBqbTiFouDI
	// https://www.youtube.com/watch?v=MlXV7qSpLDY
	//https://giangtester.com/bai-23-doc-du-lieu-tu-file-excel/
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
				driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
				driver.findElement(ByName.name("username")).sendKeys(list.get(i).getUsername());
				driver.findElement(ByName.name("pass")).click();
				;
				driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
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

	// test case ADD
	@Test(groups = "selenium", priority = 2)
	public void Test1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		try {
			driver.navigate().to("http://localhost:8008/Trung_Assignment_KTNC/login.htm");
			driver.findElement(ByName.name("username")).click();
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			driver.findElement(ByName.name("username")).sendKeys("admin");
			driver.findElement(ByName.name("pass")).click();
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			driver.findElement(ByName.name("pass")).sendKeys("admin");
			driver.findElement(ByName.name("login")).click();
			String url = "http://localhost:8008/Trung_Assignment_KTNC/user/";
			List<users> list = li.list(1);

			for (int i = 0; i < list.size(); i++) {

				driver.get(url);
				driver.findElement(ByName.name("username")).click();
				;
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				driver.findElement(ByName.name("username")).sendKeys(list.get(i).getUsername());
				driver.findElement(ByName.name("password")).click();
				;
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				driver.findElement(ByName.name("password")).sendKeys(list.get(i).getPassword());
				;
				;

				driver.findElement(ByName.name("fullname")).click();
				;
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				driver.findElement(ByName.name("fullname")).sendKeys(list.get(i).getFullname());
				;
				;
				driver.findElement(ByName.name("email")).click();
				;
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				driver.findElement(ByName.name("email")).sendKeys(list.get(i).getEmail());
				;
				;
				driver.findElement(ByName.name("admin")).click();
				;
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				if (list.get(i).isAdmin() == true) {
					driver.findElement(ByName.xpath("//*[@id=\"admin\"]/option[1]")).click();
					;
				} else {
					driver.findElement(ByName.xpath("//*[@id=\"admin\"]/option[2]")).click();
					;
				}
				driver.findElement(ByName.name("save")).click();
				;
				driver.manage().timeouts().implicitlyWait(525, TimeUnit.SECONDS);
				Thread.sleep(300);
				// *[@id="dataTable"]/tbody/tr[3]/td[1]/b/a
				// Ket mong muon thong

				// ket qua mong muon dang nhap thanh cong
				List<users> count = li.list(0);
				String error = driver.findElement(ByName.xpath("/html/body/div[2]/span[3]")).getText();
				System.out.println(error);
				if (error.equals("Vui lòng kiểm tra các lỗi nhập liệu")) {
					TestNGResults.put("2" + (count.size() + i), new Object[] { "Test_createAcount_" + (i + 1),
							"Add an account with username:" + list.get(i).getUsername() + "\t\n" + "password:"
									+ list.get(i).getPassword() + "\t\n" + "fullname:" + list.get(i).getFullname()
									+ "\t\n" + "email:" + list.get(i).getEmail() + "\t\n" + "admin:"
									+ list.get(i).isAdmin() + "\t\n",
							"Thêm Thành công", error, "FAILDED" });
					demFail.add("Faild");

				} else {
					TestNGResults.put("2" + (count.size() + i), new Object[] { "Test_createAcount_" + (i + 1),
							" Add an account with username:" + list.get(i).getUsername() + "\t\n" + "password:"
									+ list.get(i).getPassword() + "\t\n" + "fullname:" + list.get(i).getFullname()
									+ "\t\n" + "email:" + list.get(i).getEmail() + "\t\n" + "admin:"
									+ list.get(i).isAdmin() + "\t\n",
							"Thêm Thành công", error, "PASSED" });
					demPass.add("PASSED");

				}
				Assert.assertTrue(true);
			}

		} catch (Exception e) {
			// TODO: handle exception

			TestNGResults.put("3", new Object[] { "Error", "Add Login", "Thêm Thành công", "" + e, "Faild" });
			Assert.assertTrue(false);
			System.out.println(e);
		} finally {
			/// html/body/div/div[1]/div[2]/ul/li[2]/a

		}

	}

//	//test tìm kiếm
//	
//	@Test(groups = "selenium",priority = 3)
//	public void Test2() throws EncryptedDocumentException, InvalidFormatException, IOException {
//		try {
//	      String url ="http://localhost:8008/Trung_Assignment_KTNC/user/";
//			
//			 
//			 List<users> list = li.list(1);
//			 
//		    for(int i =0; i<list.size();i++) {
//				
//		    	 driver.get(url);	
//				 driver.findElement(ByName.xpath("/html/body/div/div[2]/div[1]/div[2]/div/div/div[2]/div/div/div[2]/label/input")).sendKeys(list.get(i).getFullname());;
//				 driver.manage().timeouts().implicitlyWait(50, TimeUnit.SECONDS);
//				 Thread.sleep(1000);
//				 ///
//				 List<users> count = li.list(0);
//				 List<users> count1 = li.list(1);
//				 int dem = (count.size()+count1.size());
//				 
//				 String error = driver.findElement(ByName.xpath("html/body/div/div[2]/div[1]/div[2]/div/div/div[2]/div/div/table/tbody/tr/td")).getText();
//				 String thongbao = "No matching records found";
//				 if(error.equals(thongbao)) {
//					 TestNGResults.put("2" + (dem + i), new Object[] {
//							  "Search_" + (i+1)," Search = " + list.get(i).getFullname() , "Tìm thông tin nhân viên", error,"FAILED"
//					  });
//					 demFail.add("Faild");
//				 }else {
//					 
//					 TestNGResults.put("2" + (dem + i), new Object[] {
//							 "Search_" + (i+1)," Search = " + list.get(i).getFullname() ,"Tìm thông tin nhân viên", error,"PASSED"
//					  });
//					 demPass.add("PASSED");
//				 }
//			 }
//		} catch (Exception e) {
//			// TODO: handle exception
//
//		}
//	}
	// test update
	@Test(groups = "selenium", priority = 4)
	public void Test3() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		try {

			String url = "http://localhost:8008/Trung_Assignment_KTNC/user/";
			// list update
			List<users> list = li.list(2);

			for (int i = 0; i < list.size(); i++) {

				if (list.get(i).getUsername().length() > 0) {
					driver.get(url);
					driver.navigate().to("http://localhost:8008/Trung_Assignment_KTNC/user/edit.htm?username="
							+ list.get(i).getUsername());
					driver.findElement(ByName.name("username")).click();
					;

					driver.findElement(ByName.name("password")).clear();
					;
					Thread.sleep(300);
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					driver.findElement(ByName.name("password")).sendKeys(list.get(i).getPassword());
					;
					;

					driver.findElement(ByName.name("fullname")).clear();
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					Thread.sleep(400);
					driver.findElement(ByName.name("fullname")).sendKeys(list.get(i).getFullname());
					;
					;
					driver.findElement(ByName.name("email")).clear();
					;
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					Thread.sleep(400);
					driver.findElement(ByName.name("email")).sendKeys(list.get(i).getEmail());
					;
					;
					driver.findElement(ByName.name("admin")).click();
					;
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					Thread.sleep(400);
					if (list.get(i).isAdmin() == true) {
						driver.findElement(ByName.xpath("//*[@id=\"admin\"]/option[1]")).click();
						;
					} else {
						driver.findElement(ByName.xpath("//*[@id=\"admin\"]/option[2]")).click();
						;
					}
					driver.findElement(ByName.name("save")).click();
					;
					driver.manage().timeouts().implicitlyWait(25, TimeUnit.SECONDS);
					Thread.sleep(300);
					// *[@id="dataTable"]/tbody/tr[3]/td[1]/b/a
					// Ket mong muon thong

					// ket qua mong muon dang nhap thanh cong
					List<users> count = li.list(0);
					List<users> count2 = li.list(1);
					int soTest = (count.size() + count2.size() * 2);
					String error = driver.findElement(ByName.xpath("/html/body/div[2]/span[3]")).getText();
					System.out.println(error);
					if (error.equals("Vui lòng kiểm tra các lỗi nhập liệu")) {
						TestNGResults.put("2" + (soTest + i),
								new Object[] { "Test_updateAcount_" + (i + 1),
										" Update an account with \t\n sername:" + list.get(i).getUsername() + "\t\n"
												+ "password:" + list.get(i).getPassword() + "\t\n" + "fullname:"
												+ list.get(i).getFullname() + "\t\n" + "email:" + list.get(i).getEmail()
												+ "\t\n" + "admin:" + list.get(i).isAdmin() + "\t\n",
										"Update Thành công", error, "FAILED" });
						demFail.add("Faild");

					} else {

						TestNGResults.put("2" + (soTest + i),
								new Object[] { "Test_updateAcount_" + (i + 1),
										"Update an account with \t\n username:" + list.get(i).getUsername() + "\t\n"
												+ "password:" + list.get(i).getPassword() + "\t\n" + "fullname:"
												+ list.get(i).getFullname() + "\t\n" + "email:" + list.get(i).getEmail()
												+ "\t\n" + "admin:" + list.get(i).isAdmin() + "\t\n",
										"Update Thành công", error, "PASSED" });
						demPass.add("PASSED");

					}

				} else {
				}
			}

		} catch (Exception e) {
			// TODO: handle exception

			TestNGResults.put("3", new Object[] { "Error", "Add Login", "Thêm Thành công", "" + e, "Faild" });
			Assert.assertTrue(false);
			System.out.println(e);
		} finally {
			/// html/body/div/div[1]/div[2]/ul/li[2]/as
		}
	}

	// delete
	@Test(groups = "selenium", priority = 5)
	public void Test4() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		try {

			String url = "http://localhost:8008/Trung_Assignment_KTNC/user/";
			// list update
			List<users> list = li.list(2);

			for (int i = 0; i < list.size(); i++) {

				if (list.get(i).getUsername().length() > 0) {
					driver.get(url);
					driver.navigate().to("http://localhost:8008/Trung_Assignment_KTNC/user/delete.htm?username="
							+ list.get(i).getUsername());
					Thread.sleep(1000);
					// ket qua mong muon dang nhap thanh cong
					List<users> count = li.list(0);
					List<users> count2 = li.list(1);
					int soTest = (count.size() + count2.size() * 2 + count.size());
					String error = driver.findElement(ByName.xpath("/html/body/div[2]/span[3]")).getText();
					System.out.println(error);
					if (error.equals("Vui lòng kiểm tra các lỗi nhập liệu")) {
						TestNGResults.put("2" + (soTest + i), new Object[] { "Test_deleteAcount_" + (i + 1),
								"Delete with " + list.get(i).getUsername(), "Xóa Thành công", error, "FAILED" });
						demFail.add("Faild");

					} else if (error.equals("Bạn không thể xóa tài khoản của chính mình!")) {
						TestNGResults.put("2" + (soTest + i), new Object[] { "Test_deleteAcount_" + (i + 1),
								"Delete with " + list.get(i).getUsername(), "Xóa Thành công", error, "FAILED" });
						demFail.add("Faild");
					} else {

						TestNGResults.put("2" + (soTest + i), new Object[] { "Test_deleteAcount_" + (i + 1),
								"Delete with " + list.get(i).getUsername(), "Xóa Thành công", error, "PASSED" });
						demPass.add("PASSED");

					}
				} else {

					driver.navigate().to("http://localhost:8008/Trung_Assignment_KTNC/user/");
				}
			}
		} catch (Exception e) {
			// TODO: handle exception

			TestNGResults.put("3", new Object[] { "Error", "Update", "Thêm Thành công", "" + e, "Faild" });
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

	// update
	@Test(groups = "selenium", priority = 8)
	public void Test6() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		try {
			String url = "http://localhost:8008/Trung_Assignment_KTNC/depart/index.htm";
			String url2 = "http://localhost:8008/Trung_Assignment_KTNC/depart/info.htm?DepartId=";
			// list update
			List<depart> list = li.list_depart(4);
			for (int i = 0; i < list.size(); i++) {
				if (list.get(i).getId().length() > 0) {
					driver.get(url);
					driver.navigate().to(url2 + list.get(i).getId());
					driver.findElement(ByName.name("tmpId")).clear();
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					driver.findElement(ByName.name("tmpId")).sendKeys(list.get(i).getId());
					driver.findElement(ByName.name("tmpName")).clear();
					driver.findElement(ByName.name("tmpName")).click();
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					driver.findElement(ByName.name("tmpName")).sendKeys(list.get(i).getName());
					driver.findElement(ByName.name("updatedepart")).click();
					Thread.sleep(1000);
					//  mong muon thong bao loi
					String error = driver.findElement(ByName.xpath("/html/body/div[2]/span[3]")).getText();
					// ket qua mong muon add depart thanh cong
					List<users> count = li.list(0);
					List<users> count2 = li.list(1);
					List<depart> count3 = li.list_depart(3);
					int soTest = (count.size() + count2.size() * 2 + count.size() + count3.size() * 2);
					if (error.equals("Cập nhật thông tin phòng ban " + list.get(i).getId() + " thành công!")) {
						TestNGResults.put("2" + (soTest + i),
								new Object[] { "Test_EditDepart_" + (i + 1),
										"Update a depart is ID : " + list.get(i).getId() + "\t\n Name : "
												+ list.get(i).getName(),
										"Muôn muốn cập nhật phòng ban thành công", error, "PASSED" });
						demPass.add("PASSED");
					} else {
						TestNGResults.put("2" + (soTest + i),
								new Object[] { "Test_EditDepart_" + (i + 1),
										"Update a depart is ID : " + list.get(i).getId() + "\t\n Name : "
												+ list.get(i).getName(),
										"Mong muốn cập nhật phòng ban thành công", error, "FAILED" });
						demFail.add("Faild");
					}

				} else {
				}
			}

		} catch (Exception e) {
			// TODO: handle exception
			TestNGResults.put("3", new Object[] { "Error", "Add Login", "Thêm Thành công", "" + e, "Faild" });
			Assert.assertTrue(false);
			System.out.println(e);
		} finally {
		}
	}

	// delete depart
	@Test(groups = "selenium", priority = 9)
	public void deleteDepart()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		try {
			// list update
			List<depart> list = li.list_depart(5);
			String url1 = "http://localhost:8008/Trung_Assignment_KTNC/depart/";
			String url2 = "http://localhost:8008/Trung_Assignment_KTNC/depart/info.htm?DepartId=";
			for (int i = 0; i < list.size(); i++) {
				/// html/body/div/div/div/h1

				if (list.get(i).getId().length() > 0) {

					driver.get(url2 + list.get(i).getId());
					driver.findElement(ByName.name("deletedepart")).click();
					Thread.sleep(800);
					driver.switchTo().alert().accept();
					Thread.sleep(800);
					// Ket mong muon thong bao loi
					String error = driver.findElement(ByName.xpath("/html/body/div[2]/span[3]")).getText();

					int count = li.list(1).size() * 2 + li.list(2).size() * 2 + li.list_depart(3).size() * 2
							+ list.size();
					// ket qua mong muon xóa depart thanh cong
					if (error.equals("Xóa phòng ban " + list.get(i).getId() + " thành công!")) {
						TestNGResults.put("2" + (count + i),
								new Object[] { "Test_DeleDepart_" + (i + 1),
										"Delete a depart with ID is " + list.get(i).getId(), "Xóa Thành công", error,
										"PASSED" });
						demPass.add("PASSED");
					} else {
						TestNGResults.put("2" + (count + i),
								new Object[] { "Test_DeleDepart_" + (i + 1),
										"Delete a depart with ID is " + list.get(i).getId(),
										"Kết quả chúng ta Xóa Thành công", error, "FAILED" });
						demFail.add("Faild");
					}

				}

			}

		} catch (Exception e) {
			// TODO: handle exception

			TestNGResults.put("3",
					new Object[] { "Error", "Delete Depart", "Kết quả chúng ta Xóa Thành công", "" + e, "Faild" });
			Assert.assertTrue(false);
			System.out.println(e);
			driver.get("http://localhost:8008/Trung_Assignment_KTNC/depart/");

		} finally {
			/// html/body/div/div[1]/div[2]/ul/li[2]/a

		}

	}

	// test case AddNV
	@Test(groups = "selenium", priority = 10)
	public void testAddNV()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		try {
			String url = "http://localhost:8008/Trung_Assignment_KTNC/staff/new.htm";
			List<Staffs> list = li.listStaff(6);
			System.out.println(list.size());
			System.out.println(list);

			for (int i = 0; i < list.size(); i++) {
				driver.get(url);
				driver.findElement(ByName.name("name")).click();
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				driver.findElement(ByName.name("name")).sendKeys(list.get(i).getName());

				driver.findElement(ByName.name("gender")).click();
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				if (list.get(i).isGender() == true) {
					driver.findElement(ByName.xpath("//*[@id=\"gender\"]/option[1]")).click();
				} else {
					driver.findElement(ByName.xpath("//*[@id=\"gender\"]/option[2]")).click();
				}

				driver.findElement(ByName.name("birthday")).click();
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				driver.findElement(ByName.name("birthday")).sendKeys(list.get(i).getBirthday());

				driver.findElement(ByName.name("email")).click();
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				driver.findElement(ByName.name("email")).sendKeys(list.get(i).getEmail());

				driver.findElement(ByName.name("phone")).click();
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				driver.findElement(ByName.name("phone")).sendKeys(list.get(i).getPhone());

				driver.findElement(ByName.name("depart.id")).click();
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				if (list.get(i).getDepartid().equals("DD")) {
					driver.findElement(ByName.xpath("//*[@id=\"depart.id\"]/option[1]")).click();
				} else if (list.get(i).getDepartid().equals("HC")) {
					driver.findElement(ByName.xpath("//*[@id=\"depart.id\"]/option[2]")).click();
				} else if (list.get(i).getDepartid().equals("IT")) {
					driver.findElement(ByName.xpath("//*[@id=\"depart.id\"]/option[3]")).click();
				} else if (list.get(i).getDepartid().equals("KD")) {
					driver.findElement(ByName.xpath("//*[@id=\"depart.id\"]/option[4]")).click();
				} else if (list.get(i).getDepartid().equals("KH")) {
					driver.findElement(ByName.xpath("//*[@id=\"depart.id\"]/option[5]")).click();
				} else if (list.get(i).getDepartid().equals("KT")) {
					driver.findElement(ByName.xpath("//*[@id=\"depart.id\"]/option[6]")).click();
				} else if (list.get(i).getDepartid().equals("NS")) {
					driver.findElement(ByName.xpath("//*[@id=\"depart.id\"]/option[7]")).click();
				}

				driver.findElement(ByName.name("salary")).click();
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				driver.findElement(ByName.name("salary")).sendKeys(list.get(i).getSalary());

				driver.findElement(ByName.name("notes")).click();
				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				driver.findElement(ByName.name("notes")).sendKeys(list.get(i).getNotes());

				driver.findElement(ByName.name("addstaff")).click();
				Thread.sleep(800);
				// ket qua mong muon dang nhap thanh cong

				String error = driver.findElement(ByName.xpath("/html/body/div[2]/span[3]")).getText();

				int count = li.list(1).size() * 2 + li.list(2).size() * 2 + li.list_depart(3).size() * 2
						+ list.size() * 2;
				// ket qua mong muon xóa depart thanh cong
				if (!error.equals("Vui lòng kiểm tra các lỗi nhập liệu")) {
					TestNGResults.put("2" + (count + i),
							new Object[] { "Test_Add_Staff_" + (i + 1),
									"Add a staff with Name is " + list.get(i).getName() + ", email : "
											+ list.get(i).getEmail() + ",...",
									"Kết quả chúng ta Add NV Thành công", error, "PASSED" });
					demPass.add("PASSED");
				} else {
					TestNGResults.put("2" + (count + i),
							new Object[] { "Test_Add_Staff_" + (i + 1),
									"Add a staff with Name is " + list.get(i).getName() + ", email : "
											+ list.get(i).getEmail() + ",...",
									"Kết quả chúng ta Add NV Thành công", error, "FAILED" });
					demFail.add("Faild");
				}
				Assert.assertTrue(true);
			}

		} catch (Exception e) {
			TestNGResults.put("3", new Object[] { "Error", "Điền thông tin ", "Thông tin đã được điền", "Fail" });
			Assert.assertFalse(false);
			System.out.println(e);
		}

	}

	// Test case updateNV
	@Test(groups = "selenium", priority = 11)
	public void testUpdateNV()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		try {

			String url = "http://localhost:8008/Trung_Assignment_KTNC/staff/";
			// list update
			List<Staffs> list = li.listStaff(7);
			System.out.println(list.size());
			System.out.println(list);
			for (int i = 0; i < list.size(); i++) {

				if (list.get(i).getId().length() > 0) {
					driver.get(url);
					driver.navigate().to("http://localhost:8008/Trung_Assignment_KTNC/staff/info.htm?StaffId="
							+ list.get(i).getId());
					driver.findElement(ByName.name("name")).clear();
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					driver.findElement(ByName.name("name")).sendKeys(list.get(i).getName());

					driver.findElement(ByName.name("gender")).click();
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					if (list.get(i).isGender() == true) {
						driver.findElement(ByName.xpath("//*[@id=\"gender\"]/option[1]")).click();
					} else {
						driver.findElement(ByName.xpath("//*[@id=\"gender\"]/option[2]")).click();
					}
					if (list.get(i).isActive() == true) {
						driver.findElement(ByName.xpath("//*[@id=\"active\"]/option[1]")).click();
					} else {
						driver.findElement(ByName.xpath("//*[@id=\"active\"]/option[2]")).click();
					}
					driver.findElement(ByName.name("birthday")).clear();
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					driver.findElement(ByName.name("birthday")).sendKeys(list.get(i).getBirthday());

					driver.findElement(ByName.name("email")).clear();
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					driver.findElement(ByName.name("email")).sendKeys(list.get(i).getEmail());

					driver.findElement(ByName.name("phone")).clear();
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					driver.findElement(ByName.name("phone")).sendKeys(list.get(i).getPhone());

					driver.findElement(ByName.name("depart.id")).click();
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					if (list.get(i).getDepartid().equals("DD")) {
						driver.findElement(ByName.xpath("//*[@id=\"depart.id\"]/option[1]")).click();
					} else if (list.get(i).getDepartid().equals("HC")) {
						driver.findElement(ByName.xpath("//*[@id=\"depart.id\"]/option[2]")).click();
					} else if (list.get(i).getDepartid().equals("IT")) {
						driver.findElement(ByName.xpath("//*[@id=\"depart.id\"]/option[3]")).click();
					} else if (list.get(i).getDepartid().equals("KD")) {
						driver.findElement(ByName.xpath("//*[@id=\"depart.id\"]/option[4]")).click();
					} else if (list.get(i).getDepartid().equals("KH")) {
						driver.findElement(ByName.xpath("//*[@id=\"depart.id\"]/option[5]")).click();
					} else if (list.get(i).getDepartid().equals("KT")) {
						driver.findElement(ByName.xpath("//*[@id=\"depart.id\"]/option[6]")).click();
					} else if (list.get(i).getDepartid().equals("NS")) {
						driver.findElement(ByName.xpath("//*[@id=\"depart.id\"]/option[7]")).click();
					}

					driver.findElement(ByName.name("salary")).clear();
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					driver.findElement(ByName.name("salary")).sendKeys(list.get(i).getSalary());

					driver.findElement(ByName.name("notes")).clear();
					driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
					driver.findElement(ByName.name("notes")).sendKeys(list.get(i).getNotes());

					driver.findElement(ByName.name("updatestaff")).click();

					driver.manage().timeouts().implicitlyWait(525, TimeUnit.SECONDS);
					Thread.sleep(300);
					// Ket mong muon thong

					// ket qua mong muon dang nhap thanh cong
					int count = li.list(1).size() * 2 + li.list(2).size() * 2 + li.list_depart(3).size() * 2
							+ list.size() * 2;
					String error = driver.findElement(ByName.xpath("/html/body/div[2]/span[3]")).getText();
					if (!error.equals("Vui lòng kiểm tra các lỗi nhập liệu")) {
						TestNGResults.put("2" + (count + i),
								new Object[] { "Test_Update_Staff_" + (i + 1),
										"Update a staff with Name is " + list.get(i).getName() + ", email : "
												+ list.get(i).getEmail() + ",...",
										"Kết quả chúng ta Edit NV Thành công", error, "PASSED" });
						demPass.add("PASSED");
					} else {
						TestNGResults.put("2" + (count + i),
								new Object[] { "Test_Update_Staff_" + (i + 1),
										"Update a staff with Name is " + list.get(i).getName() + ", email : "
												+ list.get(i).getEmail() + ",...",
										"Kết quả chúng ta Edit NV Thành công", error, "FAILED" });
						demFail.add("Faild");
					}
				}
			}
		} catch (Exception e) {
			TestNGResults.put("3",
					new Object[] { "Error", "Cập nhật nhân viên", "Cập nhật Thành công", "" + e, "Faild" });
			Assert.assertFalse(false);
			System.out.println(e);
		}
	}

	// Write
	@Test(groups = "selenium", priority = 13)
	public void suiteTearDown() throws EncryptedDocumentException, InvalidFormatException, IOException {
		for (int i = 0; i < demPass.size(); i++) {
			System.out.println(demPass.get(i));
		}
		// dong nay dung de tong hop bao nhieu truong hop lam duoc, bao nhiu pass, bao nhiu faile
		TestNGResults.put("3", new Object[] { "TOTAL: " + (demFail.size() + demPass.size()),
				"PASSED: " + demPass.size(), "FAILED: " + demFail.size() + "" });
		Set<String> keyset = TestNGResults.keySet();
		int rownum = 1;
		for (String key : keyset) {
			Row row = sheer.createRow(rownum++);
			Object[] objArr = TestNGResults.get(key);
			// kich thuoc cac o trong exel
			sheer.setColumnWidth(0, 9080);
			sheer.setColumnWidth(1, 20000);
			sheer.setColumnWidth(2, 15000);
			sheer.setColumnWidth(3, 10000);
			sheer.setColumnWidth(4, 8000);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				CellStyle style = workbook.createCellStyle();
				org.apache.poi.ss.usermodel.Font font = workbook.createFont();
				style.setWrapText(true);
				font.setBold(true);
				style.setFont(font);
				style.setVerticalAlignment(VerticalAlignment.CENTER);
				style.setAlignment(HorizontalAlignment.CENTER);
				style.setBorderRight(BorderStyle.THIN);
				style.setRightBorderColor(IndexedColors.BLACK.getIndex());
				style.setBorderLeft(BorderStyle.THIN);
				style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
				style.setBorderTop(BorderStyle.THIN);
				style.setTopBorderColor(IndexedColors.BLACK.getIndex());
				style.setBorderBottom(BorderStyle.THIN);
				style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
				style.setBorderBottom(BorderStyle.THIN);
				style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
				if (obj instanceof Date) {
					cell.setCellValue((Date) obj);
				} else if (obj instanceof Boolean) {
					cell.setCellValue((Boolean) obj);
				} else if (obj instanceof String) {
					cell.setCellValue((String) obj);
				} else if (obj instanceof Double) {
					cell.setCellValue((Double) obj);
				}
				cell.setCellStyle(style);

			}
		}
		// ghi file ra
		try {
			FileOutputStream out = new FileOutputStream(new File("D:\\kiem thu nang cao\\TestCase.xlsx"));
			workbook.write(out);
			Desktop.getDesktop().open(new File("D:\\kiem thu nang cao\\TestCase.xlsx"));
			out.close();
			System.out.println("xuat ra file exel thanh cong");
			driver.close();
		} catch (Exception e) {
			// TODO: handle exception
			System.out.println("That bai " + e);
		}
	}
}
