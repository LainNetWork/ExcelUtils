package lain.cn.excel.main;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;

import lain.cn.excel.common.Common;


public class ExcelRow {
	private int lineNum;//行号
	private int cellNum;//每行Cell个数，如果此个数不一致，将无法加入表，并加入异常列
	private ArrayList<Cell> cellList;
	
	public ExcelRow() {
		this.cellNum = 0;
		this.cellList = new ArrayList<>();
	}
		
	public int getLineNum() {
		return lineNum;
	}

	public void setLineNum(int lineNum) {
		this.lineNum = lineNum;
	}

	public void AddCell(Cell cell) {
		this.cellList.add(cell);
		this.cellNum++;
	}
	
	public ArrayList<Cell> getCellList() {
		return cellList;
	}

	public long getCellNum() {
		return cellNum;
	}
	
	public List<String> getRowData(){
		List<String> list = new ArrayList<>();
		for(Cell cell:this.getCellList()) {
			list.add(Common.titleTypeJudge(cell));			
		}
		return list;
	}
	@Override
	public String toString() {
		return "ExcelRow [lineNum=" + lineNum + ", cellNum=" + cellNum + ", cellList=" + cellList + "]";
	}
	
}
