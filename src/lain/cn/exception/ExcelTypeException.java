package lain.cn.exception;

import lain.cn.main.ExcelRow;

public class ExcelTypeException extends RuntimeException{
	private static final long serialVersionUID = 1L;
	private ExcelRow row;

	public ExcelTypeException(String message) {
		super(message);
	}
	
	public ExcelTypeException(String message,ExcelRow row) {
		this.row = row;
	}

	public ExcelRow getRow() {
		return row;
	}
	
	
	
}
