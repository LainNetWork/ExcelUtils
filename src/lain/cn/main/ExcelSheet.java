package lain.cn.main;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;

import lain.cn.exception.ExcelTypeException;

public class ExcelSheet {
	private String sheetName;
	private ExcelRow title;
	private List<ExcelRow> rows;
	
	public ExcelSheet(ExcelRow title) {
		this.title = title;
		this.rows = new ArrayList<ExcelRow>();
	}
	
	public boolean addRow(ExcelRow row) {
		if(row.getCellNum()!=title.getCellNum()) {
			return false;
		}else {
			//row 类型校验，以第一行数据为标准
			if(typeJudge(row)) {
				this.rows.add(row);
			}			
		}
		return true;		
	}
	
	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	
	/**
	 * 类型校验，
	 * @param row
	 * @return
	 */
	private boolean typeJudge(ExcelRow row) {
		if(rows.size()>=1) {	
			ExcelRow head = rows.get(0);
			boolean flag = true;
			for(int i = 0;i<head.getCellList().size();i++) {
				Cell cell = row.getCellList().get(i);
				Cell hCell = head.getCellList().get(i);
				if(!cell.getCellType().equals(hCell.getCellType())) {
					flag = false;
					throw new ExcelTypeException("类型不匹配！", row);
				};
			}			
			return flag;
		}else {
			return true;
		}	
	}

	public List<ExcelRow> getRows() {
		return rows;
	}
	
	
}
