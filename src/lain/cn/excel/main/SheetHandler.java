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
	private int startRow;//记录行开始位置
	private int startCol;//记录列开始位置

	
	public SheetHandler(Sheet sheet) {
		this.errorList = new ArrayList<>();
		parse(sheet);
	}
	
	private void parse(Sheet sheet) {
		String sheetName = sheet.getSheetName();		
		//获取表头，如果表头有空行，向上抛出异常		
		ExcelRow excelTitleRow = getTitle(sheet);
		this.sheet = new ExcelSheet(excelTitleRow);
		//获取数据
		getData(sheet);
		this.sheet.setSheetName(sheetName);
		
	}
	
	private ExcelRow getTitle(Sheet sheet) {
		for(int i = sheet.getFirstRowNum();i<sheet.getLastRowNum();i++) {
			Row row = sheet.getRow(i);
			if(!ifRowEmpty(row)) {
				this.setStartRow(i);
				return titleHandle(row);
			}
		}
		return null;
		
	}
	
	private boolean ifRowEmpty(Row row) {
		boolean flag = true;
		if(row == null) {
			return flag;
		}
		for(int i =row.getFirstCellNum();i<row.getLastCellNum();i++) {
			Cell cell = row.getCell(i);
			String value = Common.titleTypeJudge(cell);
			if(!"".equals(value)) {
				flag = false;
			}
		}
		return flag;
	}
	
	
	private ExcelRow titleHandle(Row row) {
		ExcelRow titleRow = new ExcelRow();
		int flag = row.getFirstCellNum();
		for(int i =row.getFirstCellNum();i<row.getLastCellNum();i++) {			
			String title = Common.titleTypeJudge(row.getCell(i));
			if(flag==i) {
				if("".equals(title)) {
					flag++;
					continue;//无视
				}
			}
			if("".equals(title)) {
				throw new ExcelHeadBlankException("表头为空！");
			}
			titleRow.AddCell(row.getCell(i));
		}		
		if(titleRow.getCellList().size()!=titleRow.getCellNum()) {
			return null;
		}		
		this.setStartCol(flag);
		return titleRow;
	}
	
	private void getData(Sheet sheet) {
		//读取数据行		
		for(int i =this.getStartRow()+1;i<=sheet.getLastRowNum();i++) {
			Row row = sheet.getRow(i);
			
			if(ifRowEmpty(row)) {
				continue;
			}
			ExcelRow excelRow = new ExcelRow();
			for(int c = this.getStartCol();c<this.sheet.getTitle().getCellNum()+this.getStartCol();c++) {
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
	
	public List<String> getTitleData(){
		
		return null;		
		
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
	
	public int getStartRow() {
		return startRow;
	}

	public void setStartRow(int startRow) {
		this.startRow = startRow;
	}

	public int getStartCol() {
		return startCol;
	}

	public void setStartCol(int startCol) {
		this.startCol = startCol;
	}
	
}
