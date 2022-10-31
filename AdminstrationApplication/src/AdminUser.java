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
public class AdminUser {

 public void AdminChoice() {
	 boolean signin=true;
	 while(signin) {
	 System.out.println("*******************************************************************");
	 System.out.println("Welcome Admin, Please select your choice as mentioned below");
	 System.out.println("1.Check and approve the user request");
	 System.out.println("2. Delete user access");
	 System.out.println("3. Generate Retailers Bill");
	 System.out.println("4. Generate Todays Sales Report");
	 System.out.println("5. Generate Sales Reports in specific dates");
	 System.out.println("6. View List of Reatilers");
	 System.out.println("7. Generate Sales report for particular retailer");
	 System.out.println("8. Generate Sales report for particular retailer between specific dates");
	 System.out.println("9. Generate Sales report for particular Brand");
	 System.out.println("10. Generate Sales report for particular Brand between specific dates");
	 System.out.println("11. Generate Sales report ProductWise");
	 System.out.println("12.Generate Sales report ProductWise between specific period");
	 System.out.println("13. Add new Retailer");
	 System.out.println("14. Delete Existing Retailer");
	 System.out.println("15. Add new Brand");
	 System.out.println("16. Delete Existing Brand");
	 System.out.println("17. View Brands");
	 System.out.println("18. Modify the existing Bill");
	 System.out.println("19. Delete new Bill");
	 System.out.println("20. Signout");
	 System.out.println("*******************************************************************");
	 System.out.println();
	 int input=inputclass.in.nextInt();
	 UserAccess u=new UserAccess();
	 inputclass.in.nextLine();
	 switch(input) {
	 case 1:
		 checkUserRequest();
		 break;
	 case 2:
		 DeleteUserAccess();
		 break;
	 case 3:
		 u.GenerateRetailerBill();
		 break;
	 case 4:
		 u.TodaySalesReport();
		 break;
	 case 5:
		 System.out.print("Please enter the From Data in (DD-MM-YYYY) format: ");
		   String from_date=inputclass.in.nextLine();
		   System.out.print("Please enter the To Data in (DD-MM-YYYY) format: ");
		   String To_date=inputclass.in.nextLine();
		   u.SpecificDateSalesReport(from_date,To_date);
		   break;
	 case 6:
		 Map<Integer,String> retailerList=u.ViewRetailerList();
		 break;
	 case 7:
		   System.out.println("Please select the retailer from the below");
		   Map<Integer,String> map=u.ViewRetailerList();
		   int key=inputclass.in.nextInt();
		   inputclass.in.nextLine();
		   String Retailer=map.get(key);
		   u.GenerateRetailerReport(Retailer);
		   break;
	 case 8:
		   System.out.println("Please select the retailer from the below");
		   Map<Integer,String> map1=u.ViewRetailerList();
		   int key1=inputclass.in.nextInt();
		   inputclass.in.nextLine();
		   String Retailer1=map1.get(key1);
		   System.out.print("Please enter the From Data in (DD-MM-YYYY) format: ");
		   String from_date1=inputclass.in.nextLine();
		   System.out.print("Please enter the To Data in (DD-MM-YYYY) format: ");
		   String To_date1=inputclass.in.nextLine();
		   u.GenerateRetailerReportSpecificdays(Retailer1,from_date1,To_date1);
		 break;
	   case 9:
		   System.out.println("Please select the Brand from the below");
		   Map<Integer,String> BrandList=u.ViewBrandsList();
		   int Selected_brand=inputclass.in.nextInt();
		   inputclass.in.nextLine();
		   String Brand=BrandList.get(Selected_brand);
		   u.GenerateBrandReport(Brand);
		   break;
	   case 10:
		   System.out.println("Please select the Brand from the below");
		   Map<Integer,String> BrandList1=u.ViewBrandsList();
		   int Selected_brand1=inputclass.in.nextInt();
		   inputclass.in.nextLine();
		   String Brand1=BrandList1.get(Selected_brand1);
		   System.out.print("Please enter the From Data in (DD-MM-YYYY) format: ");
		   String from_date_brand=inputclass.in.nextLine();
		   System.out.print("Please enter the To Data in (DD-MM-YYYY) format: ");
		   String To_date_brand=inputclass.in.nextLine();
		   u.GenerateBrandReportSpecificdays(Brand1,from_date_brand,To_date_brand);
		   break;
	   case 11:
		   u.GenerateProductWiseSalesReport();
		   break;
	   case 12:
		   System.out.print("Please enter the From Date in (DD-MM-YYYY) format: ");
		   String from_date_product=inputclass.in.nextLine();
		   System.out.print("Please enter the To Date in (DD-MM-YYYY) format: ");
		   String To_date_product=inputclass.in.nextLine();
		   u.GenerateProductWiseSpecifiedDatesSalesReport(from_date_product,To_date_product);
		   break;
	   case 13:
		   System.out.print("Please enter the new Retailer name that need to be added: ");
		   String Retailer_name=inputclass.in.nextLine();
		   AddnewRetailer(Retailer_name);
		   break;
	   case 14:
		   System.out.print("Please select the retailer from the below list: ");
		   Map<Integer,String> retailerList1=u.ViewRetailerList();
		   int Retailerkey=inputclass.in.nextInt();
		   inputclass.in.nextLine();
		   DeleteReatiler(Retailerkey);
		   break;
	   case 15:
		   System.out.print("Please enter the new Brand name that need to be added : ");
		   String Brand_name=inputclass.in.nextLine();
		   AddnewBrand(Brand_name);
		   break;
	   case 16:
		   System.out.print("Please select the Brand from the below list: ");
		   Map<Integer,String> Brand_List1=u.ViewBrandsList();
		   int Brandkey=inputclass.in.nextInt();
		   inputclass.in.nextLine();
		   DeleteBrand(Brandkey);
		   break;
	   case 17:
		   Map<Integer,String> Brand_List2=u.ViewBrandsList();
		   break;
	   case 18:
		   break;
	   case 19:
		   break;
	   case 20:
		 signin=false;
		 break;
	 } 
	 }
 }
 
 public void DeleteBrand(int key) {
	 File path=new File("C:\\Users\\ASUS\\Desktop\\project\\BrandsList.xlsx");
	   try {
		   FileInputStream stream=new FileInputStream(path);
		   Workbook book=WorkbookFactory.create(stream);
		   Sheet sheet=book.getSheetAt(0);
		   int lastind=sheet.getLastRowNum();
		   if(key!=lastind) {
			   sheet.shiftRows(key+1, lastind, -1);
		   }
		   else {
			   sheet.removeRow(sheet.getRow(lastind));
		   }
		   FileOutputStream output=new FileOutputStream(path);
		   book.write(output);
		   stream.close();
		   output.close();
		   System.out.println("Selected Brand is detailed from the list");
	   }
	   catch(Exception e) {
			 System.out.println(e);
		 }
 }
 
 public void AddnewBrand(String brand) {
	 File path=new File("C:\\Users\\ASUS\\Desktop\\project\\BrandsList.xlsx");
	   try {
		   FileInputStream stream=new FileInputStream(path);
		   Workbook book=WorkbookFactory.create(stream);
		   Sheet sheet=book.getSheetAt(0);
		   int lastind=sheet.getLastRowNum();
		   Row row=sheet.getRow(lastind);
		   int id=(int)row.getCell(0).getNumericCellValue();
		   sheet.createRow(lastind+1);
		   row=sheet.getRow(lastind+1);
		   row.createCell(0);
		   Cell cell=row.getCell(0);
		   cell.setCellValue(id+1);
		   row.createCell(1);
		   cell=row.getCell(1);
		   cell.setCellValue(brand);
		   FileOutputStream output=new FileOutputStream(path);
		   book.write(output);
		   stream.close();
		   output.close();
		   System.out.println("Congrats for new Brand member. "+brand+" added to the list");
	   }
	   catch(Exception e) {
			 System.out.println(e);
		 }
 }
 
 public void DeleteReatiler(int key) {
	 File path=new File("C:\\Users\\ASUS\\Desktop\\project\\RetailersList.xlsx");
	   try {
		   FileInputStream stream=new FileInputStream(path);
		   Workbook book=WorkbookFactory.create(stream);
		   Sheet sheet=book.getSheetAt(0);
		   int lastind=sheet.getLastRowNum();
		   if(key!=lastind) {
			   sheet.shiftRows(key+1, lastind, -1);
		   }
		   else {
			   sheet.removeRow(sheet.getRow(lastind));
		   }
		   FileOutputStream output=new FileOutputStream(path);
		   book.write(output);
		   stream.close();
		   output.close();
		   System.out.println("Selected Reatiler is detailed from the list");
	   }
	   catch(Exception e) {
			 System.out.println(e);
		 }
 }
 
 public void AddnewRetailer(String Retailer_name) {
	 File path=new File("C:\\Users\\ASUS\\Desktop\\project\\RetailersList.xlsx");
	   try {
		   FileInputStream stream=new FileInputStream(path);
		   Workbook book=WorkbookFactory.create(stream);
		   Sheet sheet=book.getSheetAt(0);
		   int lastind=sheet.getLastRowNum();
		   Row row=sheet.getRow(lastind);
		   int id=(int)row.getCell(0).getNumericCellValue();
		   sheet.createRow(lastind+1);
		   row=sheet.getRow(lastind+1);
		   row.createCell(0);
		   Cell cell=row.getCell(0);
		   cell.setCellValue(id+1);
		   row.createCell(1);
		   cell=row.getCell(1);
		   cell.setCellValue(Retailer_name);
		   FileOutputStream output=new FileOutputStream(path);
		   book.write(output);
		   stream.close();
		   output.close();
		   System.out.println("Congrats for new retailer member. "+Retailer_name+" added to the list");
	   }
	   catch(Exception e) {
			 System.out.println(e);
		 }
 }
 
 public void DeleteUserAccess() {
	 System.out.print("Please enter the user email Id:");
	 String Email = inputclass.in.nextLine();
	 try {
		 File path=new File("C:\\Users\\ASUS\\Desktop\\UserDetails.xlsx");
		 FileInputStream input=new FileInputStream(path);
		 Workbook book=WorkbookFactory.create(input);
		 Sheet sheet=book.getSheetAt(0);
		 int lastindex=sheet.getLastRowNum();
		 Boolean Access=false;
		 for(int i=1;i<=lastindex;i++) {
			 Row row=sheet.getRow(i);
			 Cell cell=row.getCell(2);
			 if(Email.equals(cell.getStringCellValue())) {
				 if(i!=lastindex) {
					 sheet.shiftRows(i+1, lastindex, -1);
				 }
				 else {
					 sheet.removeRow(sheet.getRow(i));
				 }
				 System.out.println("User Access is deactivated");
				 Access=true;
				 break;
			 }
		 }
		 input.close();
         FileOutputStream outputStream = new FileOutputStream(path);
         book.write(outputStream);
         book.close();
         outputStream.close();
		 if(!Access) {
			 System.out.println("Invalid Details. Please check the details provided");
		 }
	 }
	 catch(Exception e) {
		 System.out.println(e);
	 }
 }
 
 public void checkUserRequest() {	 
	 try {		 
	   File f=new File("C:\\Users\\ASUS\\Desktop\\UsersRequest.xlsx");
	   FileInputStream fi=new FileInputStream(f);
	   XSSFWorkbook wb = new XSSFWorkbook(fi);   
	   XSSFSheet sheet = wb.getSheetAt(0);       
	   Iterator<Row> itr = sheet.iterator();   
	   
	   int rolnum=1;
	   System.out.println("*******************************************************************");
	   while (itr.hasNext())               {  	 
	     Row row = itr.next();  
	     Iterator<Cell> cellIterator = row.cellIterator();
	     if(rolnum>1) {
	     System.out.println("User Request "+(rolnum-1)+":");
	     int colnum=0;//iterating over each column
	     String name="",email="",password="";
	     while (cellIterator.hasNext())   
	     {  
	       Cell cell = cellIterator.next(); 
	        if(colnum==0) {
	          System.out.print("UserName:"+cell.getStringCellValue()+" ");
	          name=cell.getStringCellValue();
	        } 
	        else if(colnum==1){
		      System.out.println("Email Id:"+cell.getStringCellValue());
		      email=cell.getStringCellValue();
	        }
	        else {
	        	password=cell.getStringCellValue();
	        }
	      colnum++;
	     }
	     System.out.println("Will you accept this request(Y/N)?");
	    String input = inputclass.in.nextLine();
	    if(input.equalsIgnoreCase("y") || input.equals("Y")) {
		   ApproveUserRequest(name,email,password);
	    }
	   }
	     rolnum++;
	   }
	   System.out.println("*******************************************************************");
	   System.out.println();
	   deleteExcelRows();
	 }
	 catch(Exception e){
		 System.out.println(e);
	 }
 }
 
 public void ApproveUserRequest(String name, String emaildId,String password) {
	 String excelFilePath = "C:\\Users\\ASUS\\Desktop\\UserDetails.xlsx";
     
     try {
         FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
         Workbook workbook = WorkbookFactory.create(inputStream);

         Sheet sheet = workbook.getSheetAt(0);

         Object[][] bookData = {
                 {name, emaildId,  password},
         };

         int rowCount = sheet.getLastRowNum();
         for (Object[] aBook : bookData) {
             Row row = sheet.createRow(++rowCount);
              
             int columnCount = 0;
              
             Cell cell = row.createCell(columnCount);
             //row.removeCell(cell);
             cell.setCellValue(rowCount);
              
             for (Object field : aBook) {
                 cell = row.createCell(++columnCount);
                 if (field instanceof String) {
                     cell.setCellValue((String) field);
                 } else if (field instanceof Integer) {
                     cell.setCellValue((Integer) field);
                 }
             }
           System.out.println("User "+name+" got the application access");
         }  
         inputStream.close();
         FileOutputStream outputStream = new FileOutputStream("C:\\Users\\ASUS\\Desktop\\UserDetails.xlsx");
         workbook.write(outputStream);
         workbook.close();
         outputStream.close();
          
     } catch (IOException | EncryptedDocumentException
             | InvalidFormatException ex) {
         ex.printStackTrace();
     }
 }
 
 public void deleteExcelRows() { 
	 File f=new File("C:\\Users\\ASUS\\Desktop\\UsersRequest.xlsx");
	 try {
         FileInputStream inputStream = new FileInputStream(f);
         Workbook workbook = WorkbookFactory.create(inputStream);

         Sheet sheet = workbook.getSheetAt(0);
         int lastInd=sheet.getLastRowNum();
         for(int i=1;i<=lastInd;i++) {
           sheet.removeRow(sheet.getRow(i));
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
