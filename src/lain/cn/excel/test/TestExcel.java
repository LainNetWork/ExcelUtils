package lain.cn.excel.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.junit.Test;

import lain.cn.excel.main.ExcelHandler;
import lain.cn.excel.main.SheetHandler;


public class TestExcel {
	@Test
	public void test() {
		try {
			ExcelHandler handler = new ExcelHandler(new FileInputStream("C:\\Users\\Administrator\\Desktop\\采购异常严重分析.xlsx"));
			SheetHandler sheetHandler = handler.getSheetHandlerByName("严重异常处理趋势");
			System.out.println(sheetHandler.getSheetData());
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
