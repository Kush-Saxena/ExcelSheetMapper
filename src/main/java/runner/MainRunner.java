package runner;

import java.io.IOException;

import util.ExcelUtility;

public class MainRunner {

	public static void main(String[] args) throws IOException {
	 System.out.println("Hello");
	 ExcelUtility util = new ExcelUtility(false);
	 util.mapExcelData();
	}
}
