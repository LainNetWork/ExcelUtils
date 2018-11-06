package lain.cn.main;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import lain.cn.exception.ExcelLoadException;

public class ExcelHandler {
	private Workbook book;
	
	
	public ExcelHandler(InputStream in) {
		try {
			book = WorkbookFactory.create(in);			
		} catch (EncryptedDocumentException | IOException e) {
			throw new ExcelLoadException();
		}
	}

	
	public SheetHandler getExcelSheetByName(String sheetName) {
		Sheet sheet = this.book.getSheet(sheetName);
		SheetHandler handler = new SheetHandler(sheet);		
		return handler;
	}
	
	
	
	
}
