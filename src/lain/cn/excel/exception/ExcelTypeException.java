package lain.cn.excel.exception;

import java.util.List;

import lain.cn.excel.main.ExcelRow;

public class ExcelTypeException extends RuntimeException{
	private static final long serialVersionUID = 1L;
	private List<ExcelRow> row;

	public ExcelTypeException(String message) {
		super(message);
	}
	
	public ExcelTypeException(String message,List<ExcelRow> row) {
		this.row = row;
	}

	public List<ExcelRow> getRow() {
		return row;
	}
	
	
	
}
