package com.sample.ExceltoMongo;

import java.math.BigDecimal;

public class UnderstandComapreTo {
	
	//(enrollmentRequestDTO.getFeePercentage().compareTo(new BigDecimal(state.getFeePercentage())) > 1)
	
	public static void main(String[] args) {

	      // create 2 BigDecimal objects
	      BigDecimal bg1, bg2;
	      float b = 20;
	      float c = 5;
	      float d = 10;
	      
	      bg1 = new BigDecimal("10.0");
	      bg2 = new BigDecimal("20");

	      //create int object
	      boolean res;

	      res = ((bg1.compareTo(new BigDecimal(b)))== -1);
	      System.out.println(res);// compare bg1 with bg2
	      res = ((bg1.compareTo(new BigDecimal(c)))== -1);
	      System.out.println(res);
	      res = ((bg1.compareTo(new BigDecimal(d)))== -1);
	      System.out.println(res);
	      String str1 = "Both values are equal ";
	      String str2 = "First Value is greater ";
	      String str3 = "Second value is greater";

//	      if( res == 0 )
//	         System.out.println( str1 );
//	      else if( res == 1 )
//	         System.out.println( str2 );
//	      else if( res == -1 )
//	         System.out.println( str3 );
	   }

}
