import java.io.File;
import java.io.FileInputStream;
import java.util.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LoginUser {
	String email,password;
   public void Home() {
	   System.out.println("*******************************************************************");
	   System.out.println("Please enter the access you have");
	   System.out.println("1.Admin");
	   System.out.println("2.User");
	   System.out.println("3.Back to home page");
	   System.out.println("*******************************************************************");
	   System.out.println();
	   int input=inputclass.in.nextInt();
	   inputclass.in.nextLine();
	   switch(input) {
	   case 1:
		   System.out.println("*******************************************************************");
		   System.out.println("Please enter your credentials as below");
		   System.out.print("1.Emailid:");
		   email=inputclass.in.nextLine();
		   System.out.print("2.Password:");
		   password=inputclass.in.nextLine();
		   if(AdminAccessCheck(email,password)) {
			   System.out.println("Admin Login Successful");
			   System.out.println("*******************************************************************");
			   System.out.println();
			   AdminUser a=new AdminUser();
			   a.AdminChoice();
		   }
		   else {
			   System.out.println("Invalid Details!!");
		   }
		   break;
	   case 2:
		   System.out.println("*******************************************************************");
		   System.out.println("Please enter your credentials as below");
		   System.out.print("1.Emailid:");
		   email=inputclass.in.nextLine();
		   System.out.print("2.Password:");
		   password=inputclass.in.nextLine();
		   if(UserAccessCheck(email,password)) {
			   System.out.println("User Login Successful");
			   System.out.println("*******************************************************************");
			   System.out.println();
			   UserAccess a=new UserAccess();
			   a.Home();
		   }
		   else if(userRequestAccessCheck(email,password)){
			   System.out.println("Your request is still not approved by Admin");
		   }
		   else {
			   System.out.println("Invalid Details");
		   }
		   break;
	   case 3:
		   break;
	   }
   }
   
   public boolean AdminAccessCheck(String email,String password) {
	   try {		 
		   File f=new File("C:\\Users\\ASUS\\Desktop\\AdminDetails.xlsx");
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
		        if(email.equals(cell.getStringCellValue().toString()) && password.equals(cellIterator.next().getStringCellValue())) {
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
   
   public boolean UserAccessCheck(String email,String password) {
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
		        if(email.equals(cell.getStringCellValue().toString()) && password.equals(cellIterator.next().getStringCellValue())) {
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
   
   public boolean userRequestAccessCheck(String email,String password) {
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
		        if(email.equals(cell.getStringCellValue().toString()) && password.equals(cellIterator.next().getStringCellValue())) {
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
}
