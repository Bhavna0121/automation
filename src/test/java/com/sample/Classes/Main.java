package com.sample.Classes;

import java.util.Calendar;

public class Main {

  public static void main(String[] args) {
	  Calendar cal = Calendar.getInstance();
	    
	  cal.getTime(); //Tue Mar 13 12:37:29 IST 2018

      // add 20 days to the calendar
      cal.add(Calendar.DATE, 20);
    //  System.out.println("20 days later: " + cal.getTime());
      cal.getTime();

      // subtract 2 months from the calendar
      cal.add(Calendar.MONTH, -4);
     /// System.out.println("2 months ago: " + cal.getTime());
      cal.getTime(); //Fri Feb 02 12:38:08 IST 2018

      // subtract 5 year from the calendar
      cal.add(Calendar.YEAR, -5);
      cal.getTime();
    //  System.out.println("5 years ago: " + cal.getTime());
  }
}
