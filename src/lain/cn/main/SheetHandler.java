package lain.cn.main;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import lain.cn.common.Common;
import lain.cn.error.ErrorEnum;
import lain.cn.exception.ExcelHeadBlankException;
import lain.cn.exception.ExcelTypeException;

public class SheetHandler {
	private ExcelSheet sheet;
	private List<ErrorEnum> errorList;
	
	public SheetHandler(Sheet sheet) {
		this.errorList = new ArrayList<>();
		parse(sheet);
		
	}
	
	private void parse(Sheet sheet) {
		String sheetName = sheet.getSheetName();		
		//获取表头，如果表头有空行，向上抛出异常
		Row titleRow = sheet.getRow(0);
		ExcelRow excelTitleRow = getTitle(titleRow);
		this.sheet = new ExcelSheet(excelTitleRow);
		//获取数据
		getData(sheet);
		this.sheet.setSheetName(sheetName);
		
	}
	
	private ExcelRow getTitle(Row row) {
		ExcelRow titleRow = new ExcelRow();
		for(int i =0;i<row.getLastCellNum();i++) {
			String title = Common.titleTypeJudge(row.getCell(i));
			if("".equals(title)) {
				throw new ExcelHeadBlankException("表头为空！");
			}
			titleRow.AddCell(row.getCell(i));
		}		
		if(titleRow.getCellList().size()!=titleRow.getCellNum()) {
			return null;
		}		
		return titleRow;
	}
	
	private void getData(Sheet sheet) {
		//读取数据行		
		System.out.println(sheet.getLastRowNum());
		for(int i =1;i<=sheet.getLastRowNum();i++) {
			Row row = sheet.getRow(i);
			ExcelRow excelRow = new ExcelRow();
			for(int c = 0;c<this.sheet.getTitle().getCellNum();c++) {
				Cell cell = row.getCell(c);
				excelRow.AddCell(cell);
			}
			try {
				this.sheet.addRow(excelRow);
			} catch (ExcelTypeException e) {
				ErrorEnum error = ErrorEnum.getErrorEnumByCode(Common.ERROR_ARGUMENT);
				error.setRowNum(i);
				error.setRow(e.getRow());
				this.errorList.add(error);
			}
			
		}
	}

	public ExcelSheet getSheet() {
		return sheet;
	}
	
	public List<List<String>> getSheetData(){
		List<ExcelRow> rows = this.sheet.getRows();
		List<List<String>> reList = new ArrayList<>();
		for(ExcelRow row:rows) {
			List<String> list = new ArrayList<>();
			for(Cell cell:row.getCellList()) {
				list.add(Common.titleTypeJudge(cell));			
			}
			reList.add(list);
		}
		return reList;
		
	} 


	public List<ErrorEnum> getErrorList() {
		return errorList;
	}
	
	
	
}
