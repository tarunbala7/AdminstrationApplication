import java.util.*;
import java.io.*;
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class RegisterUser {
	
	public String name,emailId,password;
	
	public void Home() {
		
		System.out.println("*******************************************************************");
		System.out.println("Please enter the below required login details");
		System.out.println("1.Name ");
		name=inputclass.in.nextLine();
		System.out.println("2.Email Id");
		emailId=inputclass.in.nextLine();
		while(checkExistingEmail(emailId) || checkUserRequestEmail(emailId) ) {
			System.out.println("This email Id already exist.Please enter another email id");
			emailId=inputclass.in.nextLine();
		}
		System.out.println("3.Password");
		password=inputclass.in.nextLine();
		AddUserRequestDetails(name,emailId,password);
		System.out.println("Successfully registered.Your requested is sent to admin!!!!!");
		System.out.println("Please note that you can the access the application only after admin approve");
		System.out.println("Please check login page for the access of application");
		System.out.println("*******************************************************************");
		System.out.println();
	}
	public boolean checkExistingEmail(String email) {
		try {		 
			   File f=new File("C:\\Users\\ASUS\\Desktop\\UserDetails.xlsx");
			   FileInputStream fi=new FileInputStream(f);
			   XSSFWorkbook wb = new XSSFWorkbook(fi);   
			   XSSFSheet sheet = wb.getSheetAt(0);       
			   Iterator<Row> itr = sheet.iterator();    
			   while (itr.hasNext())               {  	 
			     Row row = itr.next();  
			     Iterator<Cell> cellIterator = row.cellIterator();
			     int colnum=0;//iterating over each column
			     while (cellIterator.hasNext())   
			     {  
			       Cell cell = cellIterator.next(); 
			        if(colnum==2 ){
			        if(email.equals(cell.getStringCellValue().toString())) {
			        	return true;
			        }
			        }
			      colnum++;
			     }
			   }
		}
		catch(Exception e) {
			System.out.println("The error is "+e);
		}
		return false;
	}
	
	public boolean checkUserRequestEmail(String email) {
		try {		 
			   File f=new File("C:\\Users\\ASUS\\Desktop\\UsersRequest.xlsx");
			   FileInputStream fi=new FileInputStream(f);
			   XSSFWorkbook wb = new XSSFWorkbook(fi);   
			   XSSFSheet sheet = wb.getSheetAt(0);       
			   Iterator<Row> itr = sheet.iterator();    
			   while (itr.hasNext())               {  	 
			     Row row = itr.next();  
			     Iterator<Cell> cellIterator = row.cellIterator();
			     int colnum=0;//iterating over each column
			     while (cellIterator.hasNext())   
			     {  
			       Cell cell = cellIterator.next(); 
			        if(colnum==1 ){
			        if(email.equals(cell.getStringCellValue().toString())) {
			        	return true;
			        }
			        }
			      colnum++;
			     }
			   }
		}
		catch(Exception e) {
			System.out.println("The error is "+e);
		}
		return false;
	}
	
	public void AddUserRequestDetails(String name,String emailId,String password) {
		String excelFilePath = "C:\\Users\\ASUS\\Desktop\\UsersRequest.xlsx";
	     
	     try {
	         FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
	         Workbook workbook = WorkbookFactory.create(inputStream);

	         Sheet sheet = workbook.getSheetAt(0);

	         Object[][] bookData = {
	                 {name, emailId,  password},
	         };

	         int rowCount = sheet.getLastRowNum();
	         System.out.println("Insertion started");
	         for (Object[] aBook : bookData) {
	             Row row = sheet.createRow(++rowCount);	              
	             int columnCount = -1;	              
	             for (Object field : aBook) {
	                Cell cell = row.createCell(++columnCount);
	                 if (field instanceof String) {
	                     cell.setCellValue((String) field);
	                 } else if (field instanceof Integer) {
	                     cell.setCellValue((Integer) field);
	                 }
	             }
	           System.out.println("Row Instered");
	         }  
	         inputStream.close();
	         FileOutputStream outputStream = new FileOutputStream("C:\\Users\\ASUS\\Desktop\\UsersRequest.xlsx");
	         workbook.write(outputStream);
	         workbook.close();
	         outputStream.close();
	}
	     catch(Exception e) {
	    	 System.out.println(e);
	     }
	}

}
