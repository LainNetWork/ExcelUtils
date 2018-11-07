package lain.cn.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class ExcelHandler {
	private Workbook book;
	
	
	public ExcelHandler(InputStream in) {
			try {
				this.book = WorkbookFactory.create(in);
			} catch (EncryptedDocumentException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}			

	}

	
	public SheetHandler getExcelSheetByName(String sheetName) {
		Sheet sheet = this.book.getSheet(sheetName);
		SheetHandler handler = new SheetHandler(sheet);		
		return handler;
	}
	
	public static void main(String[] args) {
		ExcelHandler handler;
		try {
			handler = new ExcelHandler(new FileInputStream(new File("D:/4.xls")));
			SheetHandler SH =handler.getExcelSheetByName("国家信息表");
			System.out.println(SH.getSheetData());
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	
}
