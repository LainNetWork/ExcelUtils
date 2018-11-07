package lain.cn.excel.main;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

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
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}			

	}
	/**
	 * 根据工作簿名查找sheet处理器
	 * @param sheetName
	 * @return
	 */
	public SheetHandler getExcelSheetByName(String sheetName) {
		Sheet sheet = this.book.getSheet(sheetName);
		SheetHandler handler = new SheetHandler(sheet);		
		return handler;
	}
	/**
	 * 获取所有的工作簿名
	 * @return
	 */
	public List<String> getSheetNames(){
		return null;		
	}
	
	
	
}
