package lain.cn.main;

import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;

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

	@Override
	public String toString() {
		return "ExcelRow [lineNum=" + lineNum + ", cellNum=" + cellNum + ", cellList=" + cellList + "]";
	}
	
}
