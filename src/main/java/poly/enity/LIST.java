package poly.enity;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;



public class LIST {
	// Thực hiện LIST kèm excel
	public List<users> list(int shett_user) throws EncryptedDocumentException, InvalidFormatException, IOException{
		List<users> list = new ArrayList<users>();
		//code lấy file cần đọc
		FileInputStream inputstream = new FileInputStream(new File("testdata.xlsx"));
		Workbook workbook = WorkbookFactory.create(inputstream);
		Sheet sheet = workbook.getSheetAt(shett_user);

	
		//định dạng cột
		DataFormatter fmt = new DataFormatter();
		Iterator<Row> iterator = sheet.iterator();
		Row fRow = iterator.next();
		Cell fCell = fRow.getCell(0);
		
		  while(iterator.hasNext()) {
	    	   Row curr = iterator.next();
	    	   users test = new users();
	    		   test.setUsername(fmt.formatCellValue(curr.getCell(0)));
                    test.setPassword(fmt.formatCellValue(curr.getCell(1)));
                    test.setFullname(fmt.formatCellValue(curr.getCell(2)));
                    test.setAdmin(Boolean.parseBoolean( fmt.formatCellValue(curr.getCell(3))));
                    test.setEmail(fmt.formatCellValue(curr.getCell(4)));
	    	   list.add(test);
	    	     
	       }
	       
		return list;
	}
	
	public List<depart> list_depart(int depart) throws EncryptedDocumentException, InvalidFormatException, IOException{
		List<depart> list = new ArrayList<depart>();
		//Luồng lất file cần đọc
		FileInputStream inputstream = new FileInputStream(new File("testdata.xlsx"));
		Workbook workbook = WorkbookFactory.create(inputstream);
		Sheet sheet = workbook.getSheetAt(depart);
		//định dạng cột
		DataFormatter fmt = new DataFormatter();
		Iterator<Row> iterator = sheet.iterator();
		Row fRow = iterator.next();
		Cell fCell = fRow.getCell(0);

		while (iterator.hasNext()) {
			Row curr = iterator.next();
			depart test = new depart();
			test.setId(fmt.formatCellValue(curr.getCell(0)));
			test.setName(fmt.formatCellValue(curr.getCell(1)));
			list.add(test);

		}

		return list;
	}
	
	//LIST Staffs
	public List<Staffs> listStaff(int numbersheet)
			throws EncryptedDocumentException, InvalidFormatException, IOException, ParseException {
		List<Staffs> list = new ArrayList<Staffs>();
		// Luồng lất file cần đọc
		FileInputStream inputstream = new FileInputStream(
				new File("testdata.xlsx"));
		Workbook workbook = WorkbookFactory.create(inputstream);
		Sheet sheet = workbook.getSheetAt(numbersheet);
		// định dạng cột
		DataFormatter fmt = new DataFormatter();
		Iterator<Row> iterator = sheet.iterator();
		Row fRow = iterator.next();
		Cell fCell = fRow.getCell(0);

		while (iterator.hasNext()) {
			Row curr = iterator.next();
			Staffs test = new Staffs();
			test.setId(fmt.formatCellValue(curr.getCell(0)));
			test.setName(fmt.formatCellValue(curr.getCell(1)));
			test.setGender(curr.getCell(2).equals("true") ? true : false);
			test.setActive(curr.getCell(3).equals("true") ? true : false);
			test.setBirthday(fmt.formatCellValue(curr.getCell(4)));
			test.setEmail(fmt.formatCellValue(curr.getCell(5)));
			test.setPhone(fmt.formatCellValue(curr.getCell(6)));
			test.setDepartid(fmt.formatCellValue(curr.getCell(7)));
			test.setSalary(fmt.formatCellValue(curr.getCell(8)));
			test.setNotes(fmt.formatCellValue(curr.getCell(10)));
			list.add(test);
		}
		return list;

	}
	
	
	public List<Records> listRecord(int numbersheet)
			throws EncryptedDocumentException, InvalidFormatException, IOException, ParseException {
		List<Records> list = new ArrayList<Records>();
		// Luồng lất file cần đọc
		FileInputStream inputstream = new FileInputStream(new File(
				"testdata.xlsx"));
		Workbook workbook = WorkbookFactory.create(inputstream);
		Sheet sheet = workbook.getSheetAt(numbersheet);
		// định dạng cột
		DataFormatter fmt = new DataFormatter();
		Iterator<Row> iterator = sheet.iterator();
		Row fRow = iterator.next();
		Cell fCell = fRow.getCell(0);

		while (iterator.hasNext()) {
			Row curr = iterator.next();
			Records test = new Records();
			int idrecord = 0; 
			if (fmt.formatCellValue(curr.getCell(0)).equals("")) {
				idrecord = 0;
			}else {
				idrecord = Integer.parseInt(fmt.formatCellValue(curr.getCell(0)));
			}
			test.setId(idrecord);
			test.setStaffid(fmt.formatCellValue(curr.getCell(1)));
			test.setType(Integer.parseInt(fmt.formatCellValue(curr.getCell(2))));
			test.setReason(fmt.formatCellValue(curr.getCell(3)));
			list.add(test);
		}
		return list;
	}
	
	
}
