package lain.cn.excel.main;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
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
	public SheetHandler getSheetHandlerByName(String sheetName) {
		Sheet sheet = this.book.getSheet(sheetName);
		SheetHandler handler = new SheetHandler(sheet);		
		return handler;
	}
	/**
	 * 获取所有的工作簿名，如果是空表，则无视
	 * @return
	 */
	public List<String> getSheetNames(){
		List<String> rList = new ArrayList<>();
		for(int i=0;i<this.book.getNumberOfSheets();i++) {
			String sheetName = this.book.getSheetName(i);
			if(!ifSheetEmpty(sheetName)) {
				rList.add(sheetName);
			}
		}
		return rList;		
	}
	/**
	 * 判断该图表是否为空表
	 * @param sheetName
	 * @return 空则返回false
	 */
	private boolean ifSheetEmpty(String sheetName) {
		SheetHandler handler = getSheetHandlerByName(sheetName);
		if(handler.getSheetData().size()==0) {
			return true;
		}else {
			return false;
		}	
	}

}
