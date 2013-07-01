package net.tugdev.java2excel.test;

import java.io.IOException;

import net.tugdev.java2excel.ExcelTabbedObjectPrinter;


public class ObjectTest {

	private String myString = "hello";
	private String myString2 = "world";
	private int myInt = 1093;
	private HelloWorld hw = new HelloWorld();
	
	public String getMyString() {
		return myString;
	}

	public void setMyString(String myString) {
		this.myString = myString;
	}

	public String getMyString2() {
		return myString2;
	}

	public void setMyString2(String myString2) {
		this.myString2 = myString2;
	}

	public int getMyInt() {
		return myInt;
	}

	public void setMyInt(int myInt) {
		this.myInt = myInt;
	}

	public static void main(String[] args) throws IOException {
		ObjectTest test = new ObjectTest();
		ExcelTabbedObjectPrinter excelPrinter = new ExcelTabbedObjectPrinter();
		excelPrinter.addObject(test);
		excelPrinter.save("test.xls");
	}
	
}
