package com.sample.Classes;

public class A {
	private MyRunnable myThreadLocal = new MyRunnable();
	public static void main(String[] args) {
		
	}
	
	String test() {
		try {
			throw new Exception();
		}
		finally {
			return "abc"; 
					}
	}
}
