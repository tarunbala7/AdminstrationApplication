import java.util.*;
import java.io.*;
import java.text.*;
import java.time.*;
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class UserAccess {
   public void Home() {
	   boolean signin=true;
	   try {
	   while(signin) {
	   System.out.println("*******************************************************************");
	   System.out.println("Hi User, Welcome to the Application");
	   System.out.println("1.Generate Retailer bill");
	   System.out.println("2. Generate Sales report for today");
	   System.out.println("3. Generate Sales report between Specific dates");
	   System.out.println("4. Generate Sales report for particular retailer");
	   System.out.println("5. Generate Sales report for particular retailer between specific dates");
	   System.out.println("6. Generate Sales report for particular Brand");
	   System.out.println("7. Generate Sales report for particular Brand between specific dates");
	   System.out.println("8. Generate Sales report ProductWise");
	   System.out.println("9. Generate Sales report ProductWise between specific period");
	   System.out.println("10. signout");
	   System.out.println("*******************************************************************");
	   int input=inputclass.in.nextInt();
	   inputclass.in.nextLine();
	   switch(input) {
	   case 1:
		   GenerateRetailerBill();
		   break;
	   case 2:
		   TodaySalesReport();
		   break;
	   case 3:
		   System.out.print("Please enter the From Data in (DD-MM-YYYY) format: ");
		   String from_date=inputclass.in.nextLine();
		   System.out.print("Please enter the To Data in (DD-MM-YYYY) format: ");
		   String To_date=inputclass.in.nextLine();
		   SpecificDateSalesReport(from_date,To_date);
		   break;
	   case 4:
		   System.out.println("Please select the retailer from the below");
		   Map<Integer,String> map=ViewRetailerList();
		   int key=inputclass.in.nextInt();
		   inputclass.in.nextLine();
		   String Retailer=map.get(key);
		   GenerateRetailerReport(Retailer);
		   break;
	   case 5:
		   System.out.println("Please select the retailer from the below");
		   Map<Integer,String> map1=ViewRetailerList();
		   int key1=inputclass.in.nextInt();
		   inputclass.in.nextLine();
		   String Retailer1=map1.get(key1);
		   System.out.print("Please enter the From Data in (DD-MM-YYYY) format: ");
		   String from_date1=inputclass.in.nextLine();
		   System.out.print("Please enter the To Data in (DD-MM-YYYY) format: ");
		   String To_date1=inputclass.in.nextLine();
		   GenerateRetailerReportSpecificdays(Retailer1,from_date1,To_date1);
		 break;
	   case 6:
		   System.out.println("Please select the Brand from the below");
		   Map<Integer,String> BrandList=ViewBrandsList();
		   int Selected_brand=inputclass.in.nextInt();
		   inputclass.in.nextLine();
		   String Brand=BrandList.get(Selected_brand);
		   GenerateBrandReport(Brand);
		   break;
	   case 7:
		   System.out.println("Please select the Brand from the below");
		   Map<Integer,String> BrandList1=ViewBrandsList();
		   int Selected_brand1=inputclass.in.nextInt();
		   inputclass.in.nextLine();
		   String Brand1=BrandList1.get(Selected_brand1);
		   System.out.print("Please enter the From Data in (DD-MM-YYYY) format: ");
		   String from_date_brand=inputclass.in.nextLine();
		   System.out.print("Please enter the To Data in (DD-MM-YYYY) format: ");
		   String To_date_brand=inputclass.in.nextLine();
		   GenerateBrandReportSpecificdays(Brand1,from_date_brand,To_date_brand);
		   break;
	   case 8:
		   GenerateProductWiseSalesReport();
		   break;
	   case 9:
		   System.out.print("Please enter the From Date in (DD-MM-YYYY) format: ");
		   String from_date_product=inputclass.in.nextLine();
		   System.out.print("Please enter the To Date in (DD-MM-YYYY) format: ");
		   String To_date_product=inputclass.in.nextLine();
		   GenerateProductWiseSpecifiedDatesSalesReport(from_date_product,To_date_product);
		   break;
	   case 10:
		   signin=false;
		   break;
	   }
	   }
	   System.out.println();
	   }
	   catch(Exception e) {
		   System.out.println(e);
	   }
   }
   
   public void GenerateProductWiseSpecifiedDatesSalesReport(String from_date,String to_date) {
	   File path1=new File("C:\\Users\\ASUS\\Desktop\\project\\DetailedSales.xlsx");
	   try {
		   FileInputStream file=new FileInputStream(path1);
		   Workbook book=WorkbookFactory.create(file);
		   Sheet sheet=book.getSheetAt(0);
		   int lastind=sheet.getLastRowNum();
		   Map<String,ArrayList<Integer>> Details=new HashMap<String,ArrayList<Integer>>();
		   for(int i=1;i<=lastind;i++) {
			   Row row=sheet.getRow(i);
			   Cell cell=row.getCell(4);
			   String Product=cell.getStringCellValue();
			   cell=row.getCell(0);
			   String Date="";
				  if(cell.getCellType()==cell.CELL_TYPE_NUMERIC) {
				   SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
				   Date =  sdf.format(cell.getDateCellValue());
				  }
				  else {
					  Date=cell.getStringCellValue();
				  }
			   Date Inputfrom = new SimpleDateFormat("dd-MM-yyyy").parse(from_date);
			   Date InputTo=new SimpleDateFormat("dd-MM-yyyy").parse(to_date);
			   Date Input =new SimpleDateFormat("dd-MM-yyyy").parse(Date);
			   if(Inputfrom.compareTo(Input) <=0 && InputTo.compareTo(Input)>=0)
			   {
				   if(Details.containsKey(Product)) {
					   ArrayList<Integer> list=Details.get(Product);
					   cell=row.getCell(5);
					   int qty= list.get(0) + (int)cell.getNumericCellValue();
					   cell=row.getCell(9);
					   int amt = list.get(1) + (int)cell.getNumericCellValue();
					   list.remove(1);
					   list.remove(0);
					   list.add(0,qty);
					   list.add(1,amt);
					   Details.put(Product,list);
				   }
				   else {
					   ArrayList<Integer> list = new ArrayList<Integer>();
					   cell=row.getCell(5);					   
					   int qty=  (int)cell.getNumericCellValue();
					   cell=row.getCell(9);
					   int amt = (int)cell.getNumericCellValue();
					   list.add(0,qty);
					   list.add(1,amt);
					   Details.put(Product,list);
				   }
			   }
			   sheet.shiftRows(i, lastind, -1);;
			   i--;
			   lastind--;
		   }
		   int i=1;
		   for(String key: Details.keySet()) {
			   ArrayList<Integer> list=Details.get(key);
			   sheet.createRow(i);
			   Row row=sheet.getRow(i);
			   row.createCell(0);
			   row.createCell(1);
			   row.createCell(2);
			   Cell cell=row.getCell(0);
			   cell.setCellValue(key);
			   cell=row.getCell(1);
			   cell.setCellValue(list.get(0));
			   cell=row.getCell(2);
			   cell.setCellValue(list.get(1));
			   i++;
		   }
		   sheet.createRow(0);
		   Row row=sheet.getRow(0);
		   row.createCell(0);
		   Cell cell=row.getCell(0);
		   cell.setCellValue("Product");
		   row.createCell(1);
		   cell=row.getCell(1);
		   cell.setCellValue("Quantity");
		   row.createCell(2);
		   cell=row.getCell(2);
		   cell.setCellValue("Total");
		   file.close();
		   FileOutputStream output=new FileOutputStream("C:\\Users\\ASUS\\Desktop\\project\\ProductWiseSpecifiedDatesSalesReport.xlsx");
		   book.write(output);
		   book.close();
		   output.close();
		   System.out.println("Product Wise Sales Report is generated");
	   }
	   catch(Exception e) {
		   System.out.println(e);
	   }
   }
   
   public void GenerateProductWiseSalesReport() {
	   File path1=new File("C:\\Users\\ASUS\\Desktop\\project\\DetailedSales.xlsx");
	   try {
		   FileInputStream file=new FileInputStream(path1);
		   Workbook book=WorkbookFactory.create(file);
		   Sheet sheet=book.getSheetAt(0);
		   int lastind=sheet.getLastRowNum();
		   Map<String,ArrayList<Integer>> Details=new HashMap<String,ArrayList<Integer>>();
		   for(int i=1;i<=lastind;i++) {
			   Row row=sheet.getRow(i);
			   Cell cell=row.getCell(4);
			   String Product=cell.getStringCellValue();
				   if(Details.containsKey(Product)) {
					   ArrayList<Integer> list=Details.get(Product);
					   cell=row.getCell(5);
					   int qty= list.get(0) + (int)cell.getNumericCellValue();
					   cell=row.getCell(9);
					   int amt = list.get(1) + (int)cell.getNumericCellValue();
					   list.remove(1);
					   list.remove(0);
					   list.add(0,qty);
					   list.add(1,amt);
					   Details.put(Product,list);
				   }
				   else {
					   ArrayList<Integer> list = new ArrayList<Integer>();
					   cell=row.getCell(5);					   
					   int qty=  (int)cell.getNumericCellValue();
					   cell=row.getCell(9);
					   int amt = (int)cell.getNumericCellValue();
					   list.add(0,qty);
					   list.add(1,amt);
					   Details.put(Product,list);
				   }
				   sheet.shiftRows(i, lastind, -1);;
				   i--;
				   lastind--;
			   }
		   int i=1;
		   for(String key: Details.keySet()) {
			   ArrayList<Integer> list=Details.get(key);
			   sheet.createRow(i);
			   Row row=sheet.getRow(i);
			   row.createCell(0);
			   row.createCell(1);
			   row.createCell(2);
			   Cell cell=row.getCell(0);
			   cell.setCellValue(key);
			   cell=row.getCell(1);
			   cell.setCellValue(list.get(0));
			   cell=row.getCell(2);
			   cell.setCellValue(list.get(1));
			   i++;
		   }
		   sheet.createRow(0);
		   Row row=sheet.getRow(0);
		   row.createCell(0);
		   Cell cell=row.getCell(0);
		   cell.setCellValue("Product");
		   row.createCell(1);
		   cell=row.getCell(1);
		   cell.setCellValue("Quantity");
		   row.createCell(2);
		   cell=row.getCell(2);
		   cell.setCellValue("Total");
		   file.close();
		   FileOutputStream output=new FileOutputStream("C:\\Users\\ASUS\\Desktop\\project\\ProductWiseSalesReport.xlsx");
		   book.write(output);
		   book.close();
		   output.close();
		   System.out.println("Product Wise Sales Report is generated");
	   }
	   catch(Exception e) {
		   System.out.println(e);
	   }
   }
   
   public void GenerateBrandReportSpecificdays(String Brand,String from_date,String To_date) {
	   File path1=new File("C:\\Users\\ASUS\\Desktop\\project\\DetailedSales.xlsx");
	   try {
		   FileInputStream file=new FileInputStream(path1);
		   Workbook book=WorkbookFactory.create(file);
		   Sheet sheet=book.getSheetAt(0);
		   int lastind=sheet.getLastRowNum();
		   Map<String,ArrayList<Integer>> Details=new HashMap<String,ArrayList<Integer>>();
		   for(int i=1;i<=lastind;i++) {
			   Row row=sheet.getRow(i);
			   Cell cell=row.getCell(3);
			   String Brand_excel=cell.getStringCellValue();
			   cell=row.getCell(0);
			   String Date="";
				  if(cell.getCellType()==cell.CELL_TYPE_NUMERIC) {
				   SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
				   Date =  sdf.format(cell.getDateCellValue());
				  }
				  else {
					  Date=cell.getStringCellValue();
				  }
			   Date Inputfrom = new SimpleDateFormat("dd-MM-yyyy").parse(from_date);
			   Date InputTo=new SimpleDateFormat("dd-MM-yyyy").parse(To_date);
			   Date Input =new SimpleDateFormat("dd-MM-yyyy").parse(Date);
			   if((Brand.equals(Brand_excel)) && (Inputfrom.compareTo(Input) <=0 && InputTo.compareTo(Input)>=0) ) {  
				   if(Details.containsKey(Date)) {
					   ArrayList<Integer> list=Details.get(Date);
					   cell=row.getCell(5);
					   int qty= list.get(0) + (int)cell.getNumericCellValue();
					   cell=row.getCell(9);
					   int amt = list.get(1) + (int)cell.getNumericCellValue();
					   list.remove(1);
					   list.remove(0);
					   list.add(0,qty);
					   list.add(1,amt);
					   Details.put(Date,list);
				   }
				   else {
					   ArrayList<Integer> list = new ArrayList<Integer>();
					   cell=row.getCell(5);					   
					   int qty=  (int)cell.getNumericCellValue();
					   cell=row.getCell(9);
					   int amt = (int)cell.getNumericCellValue();
					   list.add(0,qty);
					   list.add(1,amt);
					   Details.put(Date,list);
				   }	   
			   }
			   sheet.shiftRows(i, lastind, -1);;
			   i--;
			   lastind--;
		   }
		   int i=1;
		   for(String key: Details.keySet()) {
			   ArrayList<Integer> list=Details.get(key);
			   sheet.createRow(i);
			   Row row=sheet.getRow(i);
			   row.createCell(0);
			   row.createCell(1);
			   row.createCell(2);
			   Cell cell=row.getCell(0);
			   cell.setCellValue(key);
			   cell=row.getCell(1);
			   cell.setCellValue(list.get(0));
			   cell=row.getCell(2);
			   cell.setCellValue(list.get(1));
			   i++;
		   }
		   sheet.createRow(0);
		   Row row=sheet.getRow(0);
		   row.createCell(0);
		   Cell cell=row.getCell(0);
		   cell.setCellValue("Date");
		   row.createCell(1);
		   cell=row.getCell(1);
		   cell.setCellValue("Quantity");
		   row.createCell(2);
		   cell=row.getCell(2);
		   cell.setCellValue("Total");
		   file.close();
		   FileOutputStream output=new FileOutputStream("C:\\Users\\ASUS\\Desktop\\project\\BrandSalesReportSpecificDates_"+Brand+".xlsx");
		   book.write(output);
		   book.close();
		   output.close();
		   System.out.println("Brand Sales Report is generated for "+Brand);
	   }
	   catch(Exception e) {
		   System.out.println(e);
	   }
   }
   
   public void GenerateBrandReport(String Brand) {
	   File path1=new File("C:\\Users\\ASUS\\Desktop\\project\\DetailedSales.xlsx");
	   try {
		   FileInputStream file=new FileInputStream(path1);
		   Workbook book=WorkbookFactory.create(file);
		   Sheet sheet=book.getSheetAt(0);
		   int lastind=sheet.getLastRowNum();
		   Map<String,ArrayList<Integer>> Details=new HashMap<String,ArrayList<Integer>>();
		   for(int i=1;i<=lastind;i++) {
			   Row row=sheet.getRow(i);
			   Cell cell=row.getCell(3);
			   String Brand_excel=cell.getStringCellValue();
			   if((Brand.equals(Brand_excel))) {
				   cell=row.getCell(0);
				   String Date="";
					  if(cell.getCellType()==cell.CELL_TYPE_NUMERIC) {
					   SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
					   Date =  sdf.format(cell.getDateCellValue());
					  }
					  else {
						  Date=cell.getStringCellValue();
					  }
				   if(Details.containsKey(Date)) {
					   ArrayList<Integer> list=Details.get(Date);
					   cell=row.getCell(5);
					   int qty= list.get(0) + (int)cell.getNumericCellValue();
					   cell=row.getCell(9);
					   int amt = list.get(1) + (int)cell.getNumericCellValue();
					   list.remove(1);
					   list.remove(0);
					   list.add(0,qty);
					   list.add(1,amt);
					   Details.put(Date,list);
				   }
				   else {
					   ArrayList<Integer> list = new ArrayList<Integer>();
					   cell=row.getCell(5);					   
					   int qty=  (int)cell.getNumericCellValue();
					   cell=row.getCell(9);
					   int amt = (int)cell.getNumericCellValue();
					   list.add(0,qty);
					   list.add(1,amt);
					   Details.put(Date,list);
				   }	   
			   }
			   sheet.shiftRows(i, lastind, -1);;
			   i--;
			   lastind--;
		   }
		   int i=1;
		   for(String key: Details.keySet()) {
			   ArrayList<Integer> list=Details.get(key);
			   sheet.createRow(i);
			   Row row=sheet.getRow(i);
			   row.createCell(0);
			   row.createCell(1);
			   row.createCell(2);
			   Cell cell=row.getCell(0);
			   cell.setCellValue(key);
			   cell=row.getCell(1);
			   cell.setCellValue(list.get(0));
			   cell=row.getCell(2);
			   cell.setCellValue(list.get(1));
			   i++;
		   }
		   sheet.createRow(0);
		   Row row=sheet.getRow(0);
		   row.createCell(0);
		   Cell cell=row.getCell(0);
		   cell.setCellValue("Date");
		   row.createCell(1);
		   cell=row.getCell(1);
		   cell.setCellValue("Quantity");
		   row.createCell(2);
		   cell=row.getCell(2);
		   cell.setCellValue("Total");
		   file.close();
		   FileOutputStream output=new FileOutputStream("C:\\Users\\ASUS\\Desktop\\project\\BrandSalesReport_"+Brand+".xlsx");
		   book.write(output);
		   book.close();
		   output.close();
		   System.out.println("Brand Sales Report is generated for "+Brand);
	   }
	   catch(Exception e) {
		   System.out.println(e);
	   }
   }
   
   public Map<Integer,String> ViewBrandsList() {
	   File path=new File("C:\\Users\\ASUS\\Desktop\\project\\BrandsList.xlsx");
	   try {
		   FileInputStream stream=new FileInputStream(path);
		   Workbook book=WorkbookFactory.create(stream);
		   Sheet sheet=book.getSheetAt(0);
		   Map<Integer,String> map=new HashMap<Integer,String>();
		   int lastind=sheet.getLastRowNum();
		   for(int i=1;i<=lastind;i++) {
			   Row row=sheet.getRow(i);
			   int key=i;
			   String Value=row.getCell(1).getStringCellValue();
			   System.out.println(key+". "+Value);
			   map.put(key,Value);
		   }
		   stream.close();
		   return map;
	   }
	   catch(Exception e) {
		   
	   }
	   return new HashMap<Integer,String>();
   }
   
   public Map<Integer,String> ViewProductsList() {
	   File path=new File("C:\\Users\\ASUS\\Desktop\\project\\ProductList.xlsx");
	   try {
		   FileInputStream stream=new FileInputStream(path);
		   Workbook book=WorkbookFactory.create(stream);
		   Sheet sheet=book.getSheetAt(0);
		   Map<Integer,String> map=new HashMap<Integer,String>();
		   int lastind=sheet.getLastRowNum();
		   for(int i=1;i<=lastind;i++) {
			   Row row=sheet.getRow(i);
			   int key=i;
			   String Value=row.getCell(1).getStringCellValue();
			   System.out.println(key+". "+Value);
			   map.put(key,Value);
		   }
		   stream.close();
		   return map;
	   }
	   catch(Exception e) {
		   
	   }
	   return new HashMap<Integer,String>();
   }
   
   public void GenerateRetailerReportSpecificdays(String Retailer,String from_date,String To_date) {
	   File path1=new File("C:\\Users\\ASUS\\Desktop\\project\\OverViewSaleList.xlsx");
	   try {
		   FileInputStream file=new FileInputStream(path1);
		   Workbook book=WorkbookFactory.create(file);
		   Sheet sheet=book.getSheetAt(0);
		   int lastind=sheet.getLastRowNum();
		   for(int i=1;i<=lastind;i++) {
			   Row row=sheet.getRow(i);
			   Cell cell=row.getCell(0);
			   String date="";
				  if(cell.getCellType()==cell.CELL_TYPE_NUMERIC) {
				   SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
				   date =  sdf.format(cell.getDateCellValue());
				  }
				  else {
					  date=cell.getStringCellValue();
				  }
			   Date Inputfrom = new SimpleDateFormat("dd-MM-yyyy").parse(from_date);
			   Date InputTo=new SimpleDateFormat("dd-MM-yyyy").parse(To_date);
			   Date Input =new SimpleDateFormat("dd-MM-yyyy").parse(date);
			   cell=row.getCell(2);
			   String Retailer_excel=cell.getStringCellValue();
			   if(!(Retailer.equals(Retailer_excel)) || !(Inputfrom.compareTo(Input) <=0 && InputTo.compareTo(Input)>=0)) {
				   if(i!=lastind) {
					   sheet.shiftRows(i+1, lastind,-1);
				   }
				   else {
					   sheet.removeRow(sheet.getRow(lastind));
				   }
				   i--;
				   lastind--;
			   }
		   }
		   file.close();
		   FileOutputStream output=new FileOutputStream("C:\\Users\\ASUS\\Desktop\\project\\RetailerSalesReportSpecifiedDates_"+Retailer+".xlsx");
		   book.write(output);
		   book.close();
		   output.close();
		   System.out.println("Retailers Sales Report of specific dates is generated for "+Retailer);
	   }
	   catch(Exception e) {
		   System.out.println(e);
	   }
   }
   
   public void GenerateRetailerReport(String Retailer) {
	   File path1=new File("C:\\Users\\ASUS\\Desktop\\project\\OverViewSaleList.xlsx");
	   try {
		   FileInputStream file=new FileInputStream(path1);
		   Workbook book=WorkbookFactory.create(file);
		   Sheet sheet=book.getSheetAt(0);
		   int lastind=sheet.getLastRowNum();
		   for(int i=1;i<=lastind;i++) {
			   Row row=sheet.getRow(i);
			   Cell cell=row.getCell(2);
			   String Retailer_excel=cell.getStringCellValue();
			   if(!(Retailer.equals(Retailer_excel))) {
				   if(i!=lastind) {
					   sheet.shiftRows(i+1, lastind,-1);
				   }
				   else {
					   sheet.removeRow(sheet.getRow(lastind));
				   }
				   i--;
				   lastind--;
			   }
		   }
		   file.close();
		   FileOutputStream output=new FileOutputStream("C:\\Users\\ASUS\\Desktop\\project\\RetailerSalesReport_"+Retailer+".xlsx");
		   book.write(output);
		   book.close();
		   output.close();
		   System.out.println("Retailers Sales Report is generated for "+Retailer);
	   }
	   catch(Exception e) {
		   System.out.println(e);
	   }
   }
   
   public Map<Integer,String> ViewRetailerList() {
	   File path=new File("C:\\Users\\ASUS\\Desktop\\project\\RetailersList.xlsx");
	   try {
		   FileInputStream stream=new FileInputStream(path);
		   Workbook book=WorkbookFactory.create(stream);
		   Sheet sheet=book.getSheetAt(0);
		   Map<Integer,String> map=new HashMap<Integer,String>();
		   int lastind=sheet.getLastRowNum();
		   
		   for(int i=1;i<=lastind;i++) {
			   Row row=sheet.getRow(i);
			   int key=i;
			   String Value=row.getCell(1).getStringCellValue();
			   System.out.println(key+". "+Value);
			   map.put(key,Value);
		   }
		   stream.close();
		   return map;
	   }
	   catch(Exception e) {
		   
	   }
	   return new HashMap<Integer,String>();
   }
   
   public void SpecificDateSalesReport(String from_date,String To_date) {
	   File path1=new File("C:\\Users\\ASUS\\Desktop\\project\\OverViewSaleList.xlsx");
	   try {
		   FileInputStream file=new FileInputStream(path1);
		   Workbook book=WorkbookFactory.create(file);
		   Sheet sheet=book.getSheetAt(0);
		   int lastind=sheet.getLastRowNum();
		   for(int i=1;i<=lastind;i++) {
			   Row row=sheet.getRow(i);
			   Cell cell=row.getCell(0);
			   System.out.println("The issue is here");
			   String date="";
			  if(cell.getCellType()==cell.CELL_TYPE_NUMERIC) {
			   SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
			   date =  sdf.format(cell.getDateCellValue());
			  }
			  else {
				  date=cell.getStringCellValue();
			  }
			   System.out.println("Nope!!");
			   Date Inputfrom = new SimpleDateFormat("dd-MM-yyyy").parse(from_date);
			   Date InputTo=new SimpleDateFormat("dd-MM-yyyy").parse(To_date);
			   Date Input =new SimpleDateFormat("dd-MM-yyyy").parse(date);
			   if(!(Inputfrom.compareTo(Input) <=0 && InputTo.compareTo(Input)>=0)) {
				   if(i!=lastind) {
					   sheet.shiftRows(i+1, lastind,-1);
				   }
				   else {
					   sheet.removeRow(sheet.getRow(lastind));
				   }
				   i--;
				   lastind--;
			   }
		   }
		   file.close();
		   FileOutputStream output=new FileOutputStream("C:\\Users\\ASUS\\Desktop\\project\\SalesinSpecifiedDates.xlsx");
		   book.write(output);
		   book.close();
		   output.close();
		   System.out.println("Sales Report in specific dates is generated. Please check in the destination");
	   }
	   catch(Exception e) {
		   System.out.println(e);
	   }
   }
   
   public void TodaySalesReport() {
	   File path1=new File("C:\\Users\\ASUS\\Desktop\\project\\DailySalesList.xlsx");
	   try {
		   FileInputStream file=new FileInputStream(path1);
		   Workbook book=WorkbookFactory.create(file);
		   Sheet sheet=book.getSheetAt(0);
		   int lastind=sheet.getLastRowNum();
		   Date d = Calendar.getInstance().getTime();  
		   DateFormat df = new SimpleDateFormat("dd-MM-yyyy");  
		   String sDate = df.format(d);
		   for(int i=1;i<=lastind;i++) {
			   Row row=sheet.getRow(i);
			   Cell cell=row.getCell(0);
			   if(!cell.getStringCellValue().equals(sDate)) { 
				   if(i!=lastind) {
					   sheet.shiftRows(i+1, lastind, -1);
				   }
				   else {
					   sheet.removeRow(sheet.getRow(sheet.getLastRowNum()));
				   }
				   i--;
				   lastind--;
			   }
		   }
		   file.close();
		   FileOutputStream output=new FileOutputStream(path1);
		   book.write(output);
		   book.close();
		   output.close();
		   System.out.println("Today sales Report is generated. Please check in the destination");
	   }
	   catch(Exception e) {
		   System.out.println(e);
	   }
   }
   public void GenerateRetailerBill() {
	   System.out.println("--------------Retailer Bill Generation-----------------");
	   boolean addproduct=true;
	   System.out.print("Please Enter the Bill no:");
	   int Billno=inputclass.in.nextInt();
	   inputclass.in.nextLine();
	   System.out.println("Select the Reatiler from the below list: ");
	   Map<Integer,String> map=ViewRetailerList();
	   int Retailer_id= inputclass.in.nextInt();
	   inputclass.in.nextLine();
	   String Retailer= map.get(Retailer_id);
	   AddBillAndReatiler(Billno,Retailer);
	   System.out.println("Please enter your choice and details as mentioned below");
	   while(addproduct) {
		   System.out.println("1. Add a product");
		   System.out.println("2. No need to add any more");
		   int input=inputclass.in.nextInt();
		   inputclass.in.nextLine();
		   switch(input) {
		   case 1:
			   System.out.println("Brand number from below list: ");
			   Map<Integer,String> map1=ViewBrandsList();
			   int Brand_id= inputclass.in.nextInt();
			   inputclass.in.nextLine();
			   String brand=map1.get(Brand_id);
			   System.out.print("Product number from below: ");
			   Map<Integer,String> map2=ViewProductsList();
			   int Product_id= inputclass.in.nextInt();
			   inputclass.in.nextLine();
			   String product=map2.get(Product_id);
			   System.out.print("Quantity: ");
			   int qty=inputclass.in.nextInt();
			   System.out.print("Cost: ");
			   int cost=inputclass.in.nextInt();
			   inputclass.in.nextLine();
			   AddCostandProduct(brand,product,qty,cost);
			   break;
		   case 2:
			   addproduct=false;
			   break;
		   }
		   
	   }
	   Calculatetotal();
	   System.out.print("Any Discount to be added(Y/N):");
	   String discount=inputclass.in.nextLine();
	   if(discount.equals("Y")) {
		   System.out.print("Enter the discount percentage provided to "+Retailer+":");
		   double percentage=inputclass.in.nextDouble();
		   AddDiscount(percentage);
	   }
	   GenerateReports();
	   System.out.println("Bill Generated Successfully. Please check at the Destination");
	   System.out.println("--------------End of Bill Generation-------------------");
   }
  
   
   public void GenerateReports() {
	   File path=new File("C:\\Users\\ASUS\\Desktop\\project\\BillGenerated.xlsx");
	   try {
		   FileInputStream file=new FileInputStream(path);
		   Workbook book=WorkbookFactory.create(file);
		   Sheet sheet=book.getSheetAt(0);
		   List<String> Brand_Name=new ArrayList<String>();
		   List<String> Product_Name=new ArrayList<String>();
		   List<Integer> Quantity=new ArrayList<Integer>();
		   List<Integer> Rate=new ArrayList<Integer>();
		   List<Integer> Amount=new ArrayList<Integer>();
		   for(int i=8;i<29;i++) {
			   Row row=sheet.getRow(i);
			   Cell cell=row.getCell(1);
			   if(cell.getCellType() != cell.CELL_TYPE_BLANK) {
				   Brand_Name.add(cell.getStringCellValue());
				   cell=row.getCell(2);
				   Product_Name.add(cell.getStringCellValue());
				   cell=row.getCell(3);
				   Quantity.add((int)cell.getNumericCellValue());
				   cell=row.getCell(4);
				   Rate.add((int)cell.getNumericCellValue());
				   cell=row.getCell(5);
				   Amount.add((int)cell.getNumericCellValue());   
			   }
			   else {
				   break;
			   }
		   }
		   Row row=sheet.getRow(5);
		   Cell cell=row.getCell(4);
		   Date d = Calendar.getInstance().getTime();  
		   DateFormat df = new SimpleDateFormat("dd-MM-yyyy");  
		   String date = df.format(d);
		   cell=row.getCell(2);
		   int bill_no=(int)cell.getNumericCellValue();
		   row=sheet.getRow(6);
		   cell=row.getCell(2);
		   String Retailer_Name = cell.getStringCellValue();
		   row=sheet.getRow(30);
		   cell=row.getCell(4);
		   double discount=0;
		   if(cell.getCellType()!=cell.CELL_TYPE_BLANK)
		   {
		   discount = cell.getNumericCellValue();
		   }
		   row=sheet.getRow(29);
		   cell=row.getCell(5);
		   int Total=(int)cell.getNumericCellValue();
		   row=sheet.getRow(32);
		   cell=row.getCell(5);
		   int GrandTotal=(int)cell.getNumericCellValue();
		   AddDataDetailedSalesReports(date,bill_no,Retailer_Name,Brand_Name,Product_Name,Quantity,Rate,Amount,discount);
		   AddDataOverViewSalesReports(date,bill_no,Retailer_Name,Total,discount,GrandTotal);
		   AddDataTodaySalesReports(date,bill_no,Retailer_Name,Total,discount,GrandTotal);
		   file.close();
	   }
	   catch(Exception e) {
		   System.out.println(e);
	   }
   }
   
   public void AddDataTodaySalesReports(String date,int bill_no,String RetailerName,int Total,double discount,int GrandTotal) {
	   File path1=new File("C:\\Users\\ASUS\\Desktop\\project\\DailySalesList.xlsx");
	   try {
		   FileInputStream file=new FileInputStream(path1);
		   Workbook book=WorkbookFactory.create(file);
		   Sheet sheet=book.getSheetAt(0);
		   int lastind=sheet.getLastRowNum();
		   sheet.createRow(lastind+1);
		   Row row=sheet.getRow(lastind+1);
		   row.createCell(0);
		   Cell cell=row.getCell(0);
		   cell.setCellValue(date);
		   row.createCell(1);
		   cell=row.getCell(1);
		   cell.setCellValue(bill_no);
		   row.createCell(2);
		   cell=row.getCell(2);
		   cell.setCellValue(RetailerName);
		   row.createCell(3);
		   cell=row.getCell(3);
		   cell.setCellValue(Total);
		   row.createCell(4);
		   cell=row.getCell(4);
		   cell.setCellValue(discount*100);
		   row.createCell(5);
		   cell=row.getCell(5);
		   cell.setCellValue(GrandTotal);
		   file.close();
		   FileOutputStream output=new FileOutputStream(path1);
		   book.write(output);
		   book.close();
		   output.close();
	   }
	   catch(Exception e) {
		   System.out.println(e);
	   }
   }
   
   public void AddDataOverViewSalesReports(String date,int bill_no,String RetailerName,int Total,double discount,int GrandTotal) {
	   File path1=new File("C:\\Users\\ASUS\\Desktop\\project\\OverViewSaleList.xlsx");
	   try {
		   FileInputStream file=new FileInputStream(path1);
		   Workbook book=WorkbookFactory.create(file);
		   Sheet sheet=book.getSheetAt(0);
		   int lastind=sheet.getLastRowNum();
		   sheet.createRow(lastind+1);
		   Row row=sheet.getRow(lastind+1);
		   row.createCell(0);
		   Cell cell=row.getCell(0);
		   cell.setCellValue(date);
		   row.createCell(1);
		   cell=row.getCell(1);
		   cell.setCellValue(bill_no);
		   row.createCell(2);
		   cell=row.getCell(2);
		   cell.setCellValue(RetailerName);
		   row.createCell(3);
		   cell=row.getCell(3);
		   cell.setCellValue(Total);
		   row.createCell(4);
		   cell=row.getCell(4);
		   cell.setCellValue(discount*100);
		   row.createCell(5);
		   cell=row.getCell(5);
		   cell.setCellValue(GrandTotal);
		   file.close();
		   FileOutputStream output=new FileOutputStream(path1);
		   book.write(output);
		   book.close();
		   output.close();
	   }
	   catch(Exception e) {
		   System.out.println(e);
	   }
   }
   
   public void AddDataDetailedSalesReports(String date,int bill_no,String Retailer_Name,List<String> Brand_Name,List<String> Product_Name,List<Integer> Quantity,List<Integer> rate,List<Integer> amount,double discount) {
	   File path=new File("C:\\Users\\ASUS\\Desktop\\project\\DetailedSales.xlsx");
	   try {
		   FileInputStream file=new FileInputStream(path);
		   Workbook book=WorkbookFactory.create(file);
		   Sheet sheet=book.getSheetAt(0);
		   int emptycell=sheet.getLastRowNum(),i=0;
		   for(i=0;i<Product_Name.size();i++) {	   
			   Row row=sheet.createRow(emptycell+i+1);
			   Cell cell=row.createCell(0);
			   cell.setCellValue(date);
			   cell=row.createCell(1);
			   cell.setCellValue(bill_no);
			   cell=row.createCell(2);
			   cell.setCellValue(Retailer_Name);
			   cell=row.createCell(3);
			   cell.setCellValue(Brand_Name.get(i));
			   cell=row.createCell(4);
			   cell.setCellValue(Product_Name.get(i));
			   cell=row.createCell(5);
			   cell.setCellValue(Quantity.get(i));
			   cell=row.createCell(6);
			   cell.setCellValue(rate.get(i));
			   cell=row.createCell(7);
			   cell.setCellValue(amount.get(i));
			   cell=row.createCell(8);
			   cell.setCellValue(discount*100);
			   cell=row.createCell(9);
			   int disamt=(int)(amount.get(i)-(amount.get(i) * discount ));
			   cell.setCellValue(disamt);
		   }
		   file.close();
		   FileOutputStream output=new FileOutputStream(path);
		   book.write(output);
		   book.close();
		   output.close();
	   }
	   catch(Exception e) {
		   System.out.println(e);
	   }    
   }
   
   public void AddDiscount(double percentage) {
	   File path=new File("C:\\Users\\ASUS\\Desktop\\project\\BillGenerated.xlsx");
	   try {
		   FileInputStream file=new FileInputStream(path);
		   Workbook Book = WorkbookFactory.create(file);
		   Sheet sheet=Book.getSheetAt(0);
		   Row row=sheet.getRow(29);
		   Cell cell=row.getCell(5);
		   int total = (int)cell.getNumericCellValue();
		   int discount=(int)((total * percentage)/100);
		   row=sheet.getRow(30);
		   cell=row.getCell(2);
		   cell.setCellValue("Discount");
		   cell=row.getCell(4);
		   cell.setCellValue(percentage/100);
		   cell=row.getCell(5);
		   cell.setCellValue(discount);
		   row=sheet.getRow(32);
		   cell=row.getCell(5);
		   cell.setCellValue(total-discount);
		   FileOutputStream output=new FileOutputStream(path);
		   Book.write(output);
		   Book.close();
		   output.close();
	   }
	   catch(Exception e) {
		   System.out.println(e);
	   }
   }
   
   public void Calculatetotal() {
	   File path=new File("C:\\Users\\ASUS\\Desktop\\project\\BillGenerated.xlsx");
	   try {
		   FileInputStream file=new FileInputStream(path);
		   Workbook Book = WorkbookFactory.create(file);
		   Sheet sheet=Book.getSheetAt(0);
		   int total = 0;
		   int quantity=0;
		   for(int i=8;i<24;i++) {
			   Row row=sheet.getRow(i);
			   Cell cell=row.getCell(5);
			   if(cell.getCellType() != cell.CELL_TYPE_BLANK) {
				   total=total+(int)cell.getNumericCellValue();
			   }
			   else {
				   break;
			   }
			   cell=row.getCell(3);
			   quantity=quantity + (int) cell.getNumericCellValue();
		   }
		   Row row=sheet.getRow(29);
		   Cell cell=row.getCell(5);
		   cell.setCellValue(total);
		   cell=row.getCell(3);
		   cell.setCellValue(quantity);
		   row=sheet.getRow(32);
		   cell=row.getCell(5);
		   cell.setCellValue(total);
		   FileOutputStream output=new FileOutputStream(path);
		   Book.write(output);
		   Book.close();
		   output.close();
	   }
	   catch(Exception e) {
		   System.out.println(e);
	   }
   }
   
   public void AddCostandProduct(String brand,String product,int qty,int cost) {
	   File path=new File("C:\\Users\\ASUS\\Desktop\\project\\BillGenerated.xlsx");
	   try {
		   FileInputStream file=new FileInputStream(path);
		   Workbook Book = WorkbookFactory.create(file);
		   Sheet sheet=Book.getSheetAt(0);
		   for(int i=7;i<24;i++) {
			   Row row=sheet.getRow(i);
			   Cell cell=row.getCell(1);
			   if(cell.getCellType()==cell.CELL_TYPE_BLANK) {
				   cell.setCellValue(brand);
				   cell=row.getCell(2);
				   cell.setCellValue(product);
				   cell=row.getCell(3);
				   cell.setCellValue(qty);
				   cell=row.getCell(4);
				   cell.setCellValue(cost);
				   cell=row.getCell(5);
				   cell.setCellValue(cost*qty);
				   break;
			   }
		   }
		   FileOutputStream output=new FileOutputStream(path);
		   Book.write(output);
		   Book.close();
		   output.close();
	   }
	   catch(Exception e) {
		   System.out.println(e);
	   }
   }
   
   public void AddBillAndReatiler(int Billno,String Retailer) {
	   File excelSheet=new File("C:\\Users\\ASUS\\Desktop\\project\\BillFormat.xlsx");
	   try {
		   FileInputStream f=new FileInputStream(excelSheet);
		   Workbook book=WorkbookFactory.create(f);
		   Sheet sheet=book.getSheetAt(0);
		   Row row=sheet.getRow(6);
		   Cell cell=row.getCell(2);
		   cell.setCellValue(Retailer);
		   row=sheet.getRow(5);
		   cell=row.getCell(2);
		   cell.setCellValue(Billno);
		   f.close();
		   FileOutputStream output=new FileOutputStream("C:\\Users\\ASUS\\Desktop\\project\\BillGenerated.xlsx");
		   book.write(output);
		   book.close();
		   output.close();
	   }
	   catch(Exception e) {
		   System.out.println(e);
	   }
   }
}
