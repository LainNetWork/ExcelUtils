package lain.cn.excel.main;

import java.util.ArrayList;
import java.util.List;

import lain.cn.excel.exception.ExcelLoadException;


public class ExcelSheet {
	private String sheetName;
	private ExcelRow title;
	private List<ExcelRow> rows;
	
	public ExcelSheet(ExcelRow title) {
		this.title = title;
		this.rows = new ArrayList<ExcelRow>();
	}
	
	public void addRow(ExcelRow row) {
		if(row.getCellNum()!=title.getCellNum()) {
			throw new ExcelLoadException("行数不匹配!");
		}else {
			this.rows.add(row);
		}
	}
	
	public String getSheetName() {
		return sheetName;
	}

	public ExcelRow getTitle() {
		return title;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	
	public List<ExcelRow> getRows() {
		return rows;
	}

	
}
