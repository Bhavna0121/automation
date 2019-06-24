package com.sample.Classes;

public class B extends A {
	String a;

	public B() {
		System.out.println("");
	}

	public B(String a) {
		this.a = a;
	}

	public B(B p1) {
		// TODO Auto-generated constructor stub
	}

	public void execute() {
		System.out.println("Class B method ");
	}

}
