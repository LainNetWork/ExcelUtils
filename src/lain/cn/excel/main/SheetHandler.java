package lain.cn.excel.main;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import lain.cn.excel.common.Common;
import lain.cn.excel.error.ErrorVO;
import lain.cn.excel.exception.ExcelHeadBlankException;
import lain.cn.excel.exception.ExcelTypeException;

public class SheetHandler {
	private ExcelSheet sheet;
	private List<ErrorVO> errorList;
	
	public SheetHandler(Sheet sheet) {
		this.errorList = new ArrayList<>();
		parse(sheet);
	}
	
	private void parse(Sheet sheet) {
		String sheetName = sheet.getSheetName();		
		//获取表头，如果表头有空行，向上抛出异常
		Row titleRow = sheet.getRow(sheet.getFirstRowNum());
		ExcelRow excelTitleRow = getTitle(titleRow);
		this.sheet = new ExcelSheet(excelTitleRow);
		//获取数据
		getData(sheet);
		this.sheet.setSheetName(sheetName);
		
	}
	
	private ExcelRow getTitle(Row row) {
		ExcelRow titleRow = new ExcelRow();
		for(int i =row.getFirstCellNum();i<row.getLastCellNum();i++) {
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
		for(int i =sheet.getFirstRowNum()+1;i<=sheet.getLastRowNum();i++) {
			Row row = sheet.getRow(i);
			ExcelRow excelRow = new ExcelRow();
			for(int c = row.getFirstCellNum();c<this.sheet.getTitle().getCellNum()+row.getFirstCellNum();c++) {
				Cell cell = row.getCell(c);
				excelRow.setLineNum(i);
				excelRow.AddCell(cell);
			}
			try {
				this.sheet.addRow(excelRow);
			} catch (ExcelTypeException e) {
				ErrorVO error = new ErrorVO(Common.ERROR_ARGUMENT, e.getMessage());
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
			reList.add(row.getRowData());
		}
		return reList;
		
	} 
	
	
	
	public List<List<String>> getErrorData(){
		List<List<String>> reList = new ArrayList<>();
		for(ErrorVO vo:this.getErrorList()) {
			ExcelRow row= vo.getRow();
			reList.add(row.getRowData());
		}
		return reList;
	}
	

	public List<ErrorVO> getErrorList() {
		return errorList;
	}
	
	
	
}
