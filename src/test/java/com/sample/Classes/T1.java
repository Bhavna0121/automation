package com.sample.Classes;

public class T1 implements Runnable {
	String str = "JAVA";

	String str2 = "ABC";

	public void run() {
		while (true) {
			synchronized (str2) {
				synchronized (str) {
					System.out.println("lock");
				}

			}
		}

	}

}
