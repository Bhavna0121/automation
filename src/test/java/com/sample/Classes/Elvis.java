package com.sample.Classes;

import java.util.HashSet;
import java.util.Set;

public class Elvis {

	// public static final Elvis ELVIS = new Elvis();
	//
	// private Elvis() {
	//
	// }
	//
	// private static final Boolean LIVING = true;
	// private final Boolean alive = LIVING;
	//
	// public final Boolean lives() {
	// return alive;
	// }

	public static void main(String[] args) {
		// System.out.println(ELVIS.lives() ? "Hound Dog" : "heartbreak Hotel");

//		Set<Short> s = new HashSet<Short>();
//		for (short i = 0; i < 100; i++) {
//			s.add(i);
//			s.remove(i - 1);
//		}
//		System.out.println(s.size()); //100

		 double funds = 1.00;
		 int itemBought = 0;
		 for (double price = 0.10; funds >= price; price += 0.10) {
		 funds -= price;
		 itemBought++;
		
		 }
		
		 System.out.println("item bought" + itemBought);
		 System.out.println("Chnage " + funds);
	}

}
