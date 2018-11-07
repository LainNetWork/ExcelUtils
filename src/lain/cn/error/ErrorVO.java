package lain.cn.error;

import lain.cn.main.ExcelRow;

public class ErrorVO {
	private int code;
	private String message;
	//记录出错的位置
	private int rowNum;//行数
	private int colNum;//列数
	private ExcelRow row;//出错的Row数据
	public ErrorVO(int code,String message) {
		this.code = code;
		this.message = message;
	}
	public String getMessage() {
		return message;
	}
	public void setMessage(String message) {
		this.message = message;
	}
	public int getCode() {
		return code;
	}
	public void setCode(int code) {
		this.code = code;
	}

	public long getRowNum() {
		return rowNum;
	}

	public void setRowNum(int rowNum) {
		this.rowNum = rowNum;
	}

	public long getColNum() {
		return colNum;
	}

	public void setColNum(int colNum) {
		this.colNum = colNum;
	}

	public ExcelRow getRow() {
		return row;
	}

	public void setRow(ExcelRow row) {
		this.row = row;
	}
	
	
	
}
