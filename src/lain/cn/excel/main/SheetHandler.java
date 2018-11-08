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
	private TypeHandler typeHandler;
	
	
	public SheetHandler(Sheet sheet) {
		this.errorList = new ArrayList<>();
		parse(sheet);
	}
	
	private void parse(Sheet sheet) {
		String sheetName = sheet.getSheetName();		
		//获取表头，如果表头有空行，向上抛出异常		
		ExcelRow excelTitleRow = getTitle(sheet);
		this.typeHandler = new TypeHandler(excelTitleRow);
		this.sheet = new ExcelSheet(excelTitleRow);
		//获取数据
		getData(sheet);
		this.sheet.setSheetName(sheetName);
		
	}
	
	private ExcelRow getTitle(Sheet sheet) {
		for(int i = sheet.getFirstRowNum();i<=sheet.getLastRowNum();i++) {
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
			titleRow.AddCell(row.getCell(i));
		}		
		if(titleRow.getCellList().size()!=titleRow.getCellNum()) {
			return null;
		}		
		this.setStartCol(flag);
		//去除尾部空行，判断中间是否有空
		emptyTitleHandle(titleRow);
		return titleRow;
	}
	
	private void emptyTitleHandle(ExcelRow titleRow) {
		List<Cell> cellList =titleRow.getCellList();
		for(int i =cellList.size()-1;i>=0;i--){
			Cell cell = cellList.get(i);
			String title = Common.titleTypeJudge(cell);
			if("".equals(title)) {
				titleRow.setCellNum(titleRow.getCellNum()-1);
				cellList.remove(i);
			}else {//一旦有不为空的，立刻跳出循环
				break;
			}
		}
		//判断是否包含空表头
		for(Cell cell :cellList) {
			String title = Common.titleTypeJudge(cell);
			if("".equals(title)) {
				throw new ExcelHeadBlankException("表头包含空单元格！");
			}
		}	
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
			this.typeHandler.vote(excelRow);//统计类型
			this.sheet.addRow(excelRow);
		}
		//过滤错误类型
		try {
			this.typeHandler.erroTypeFilter(this.sheet);
		} catch (ExcelTypeException e) {
			for(ExcelRow row :e.getRow()) {
				ErrorVO vo =new ErrorVO(Common.ERROR_ARGUMENT, "类型不匹配!");
				vo.setRow(row);
				this.errorList.add(vo);
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
	
	
	public List<String> getTypes(){
		return this.typeHandler.getResult();
	}

	public List<String> getTitleData(){
		
		return this.sheet.getTitle().getRowData();		
		
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

	private void setStartRow(int startRow) {
		this.startRow = startRow;
	}

	public int getStartCol() {
		return startCol;
	}

	private void setStartCol(int startCol) {
		this.startCol = startCol;
	}
	
}
