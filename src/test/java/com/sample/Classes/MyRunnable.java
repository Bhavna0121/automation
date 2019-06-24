package com.sample.Classes;

public class MyRunnable implements Runnable {

	private ThreadLocal<Integer> threadLocal = new ThreadLocal<Integer>();

	public static void main(String[] args) throws InterruptedException {
		MyRunnable myRunnable = new MyRunnable();
		Thread thread1 = new Thread(myRunnable);
		Thread thread2 = new Thread(myRunnable);
		thread1.start();
		thread2.start();
		thread1.join();
		thread2.join();
	}

	public void run() {
		threadLocal.set((int) (Math.random() * 100D));
		try {
			Thread.sleep(2000);
		} catch (InterruptedException e) {
		}

		System.out.println(threadLocal.get());
	}

}
