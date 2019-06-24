package com.sample.Classes;

public class Geeks implements Runnable {

	public void run() {
		Lock();

	}

	void Lock() {
		// System.out.println(Thread.currentThread().getName());
		synchronized (this) {
			System.out.println("in block " + Thread.currentThread().getName());
			System.out.println("in block " + Thread.currentThread().getName() + " end");
		}
	}

	public static void main(String[] args) {
		Geeks r = new Geeks();
		Thread t1 = new Thread(r);
		Thread t2 = new Thread(r);
		Geeks r1 = new Geeks();
		Thread t3 = new Thread(r1);
		t1.setName("t1");
		t2.setName("t2");
		t3.setName("t3");
		t1.start();
		t2.start();
		t3.start();
	}
}
