package com.sample.Classes;

public class HeavyWorkRunnable implements Runnable {

	public void run() {
		System.out.println("Doing heavy processing - START" + Thread.currentThread().getName());
		try {
			Thread.sleep(1000);
			doDBProcessing();
		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println("Doing heavy processing -END " + Thread.currentThread().getName());
	}

	private void doDBProcessing() throws InterruptedException {
		Thread.sleep(5000);

	}

}
