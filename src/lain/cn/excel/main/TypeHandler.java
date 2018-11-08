package lain.cn.excel.main;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import lain.cn.excel.exception.ExcelTypeException;

public class TypeHandler {
	private Map<String, Integer[]> voteMap;
	private static final String TEXT_PARTTEN = "[0-9]{1,4}-[0-9]{1,2}-[0-9]{1,2} [0-9]{1,2}:[0-9]{1,2}:[0-9]{1,2}";
	private static final String NUM = "NUM";
	private static final String TEXT = "TEXT"; 	
	private static final String DATE = "DATE";
	private List<String> col;
	public TypeHandler(ExcelRow titleRow) {
		this.col = titleRow.getRowData();
		this.voteMap = new HashMap<String, Integer[]>(3);
		this.voteMap.put(NUM, new Integer[titleRow.getCellNum()]);
		this.voteMap.put(TEXT, new Integer[titleRow.getCellNum()]);
		this.voteMap.put(DATE, new Integer[titleRow.getCellNum()]);
	}
	
	//类型决断
	public void vote(ExcelRow row) {
		List<String> rows = row.getRowData();
		for(int i = 0;i<rows.size();i++) {
			String type = typeJudge(rows.get(i));
			Integer[] list = voteMap.get(type);
			if(list[i]==null) {
				list[i] = 0;
			}
			list[i]=list[i]+1;
			voteMap.put(type, list);
		}
	}
	
	private String typeJudge(String value) {
		//是否是数字
		if(isDouble(value)) {
			return NUM;
		}
		//是否是日期
		if(Pattern.matches(TEXT_PARTTEN, value)) {
			return DATE;
		}
		return TEXT;		
	}
	
	private boolean isDouble(String value) {
		try {
			Double.parseDouble(value);
		} catch (Exception e) {
			return false;
		}
		return true;
	}
	
	//返回类型映射
	public List<String> getResult(){
		List<String> list = new ArrayList<String>();
		Integer[] numList = this.voteMap.get(NUM);
		Integer[]  dateList = this.voteMap.get(DATE);
		Integer[]  textList = this.voteMap.get(TEXT);
		for(int i = 0;i<this.col.size();i++) {			
			int countNum = numList[i]==null?0:numList[i];
			int countDate = dateList[i]==null?0:dateList[i];
			int countText = textList[i]==null?0:textList[i];
			if(countText >= countNum) {
				if(countText>=countDate) {
					
					list.add(TEXT);
				}else {
					list.add(DATE);
				}
			}else if(countNum>=countDate){
				list.add(NUM);
			}
			
		}
		return list;
	}
	
	//过滤错误行，遇到错误行时抛出异常
	public void erroTypeFilter(ExcelSheet sheet) {
		List<ExcelRow> list = sheet.getRows();
		Iterator<ExcelRow> it = list.iterator();
		List<ExcelRow> errorList =new ArrayList<>();
		while(it.hasNext()) {
			ExcelRow row = it.next();
			if(!isTypeMatch(row)) {			
				ExcelRow tempRow = new ExcelRow();
				tempRow.setCellList(row.getCellList());
				tempRow.setCellNum(row.getCellNum());
				tempRow.setLineNum(row.getLineNum());
				errorList.add(tempRow);
				it.remove();				
			}
		}
		if(errorList.size()>0) {
			throw new ExcelTypeException("类型不一致！",errorList);
		}
		
	}
	private boolean isTypeMatch(ExcelRow row) {
		List<String> list = getResult();
		List<String> rowList = row.getRowData();
		for(int i=0;i<list.size();i++) {
			String originType = list.get(i);
			if(!originType.equals(typeJudge(rowList.get(i)))) {//类型不匹配
				return false;
			}
		}
		return true;
		
	}
	
}
