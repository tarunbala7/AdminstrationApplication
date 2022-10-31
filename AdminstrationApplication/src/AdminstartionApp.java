
import java.util.*;

public class Startup {
	 
	public static void main(String[] args) { 
	  System.out.println("Welcome to Sales Adminstartion of Anuj Agencies");
	  System.out.println();
	  boolean exitstatus=true;
	  while(exitstatus) {
	  System.out.println("*******************************************************************");
	  System.out.println("Before to proceed further,Please select as mentioned below to login");
	  System.out.println("1. Regsiter ");
	  System.out.println("2. Login ");
	  System.out.println("3. Exit");
	  System.out.println("*******************************************************************");
	  System.out.println();
	  int l=1;
	     l=inputclass.in.nextInt();
	     inputclass.in.nextLine();
	  switch(l) {
	  case 1:
		  RegisterUser r=new RegisterUser();
		  r.Home();
	  case 2:
		  LoginUser L=new LoginUser();
		  L.Home();
		  break;
	  case 3:
		  exitstatus=false;
		  break;
	  }
	  }
	  inputclass.in.close();
	 System.out.println("Hope am useful.See you soon!!!!!!");
   }
}
