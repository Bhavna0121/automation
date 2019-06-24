package com.sample.Classes;

public class T2 implements Runnable {
	String str = "JAVA";

	String str2 = "ABC";

	public void run() {
		while (true) {
			synchronized (str) {
				synchronized (str2) {
					System.out.println("loc");
				}

			}
		}
	}

}
