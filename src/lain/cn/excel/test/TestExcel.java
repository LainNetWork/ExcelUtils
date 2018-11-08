package lain.cn.excel.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.junit.Test;

import lain.cn.excel.main.ExcelHandler;
import lain.cn.excel.main.SheetHandler;


public class TestExcel {
	@Test
	public void test() {
		try {
			ExcelHandler handler = new ExcelHandler(new FileInputStream("D:/Lain.xlsx"));
			SheetHandler sheetHandler = handler.getSheetHandlerByName("员工工资表");
			System.out.println(sheetHandler.getErrorData());
			System.out.println(sheetHandler.getTypes());
			System.out.println(sheetHandler.getSheetData());
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	@Test
	public void test2() {
		try {
			Date fa = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").parse("-66-22 88:66:7");

			System.out.println(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(fa));
		} catch (ParseException e) {
			e.printStackTrace();
		}
		
	}
	

}
