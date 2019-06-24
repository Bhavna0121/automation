package com.sample.Classes;

public class MainShape {

	public static void main(String[] args) {
		MainShape mainShape = new MainShape();
		Shape triangle = new Triangle();
		mainShape.drawShape(triangle);
	}

	public void drawShape(Shape shape) {
		shape.draw();
	}

}
