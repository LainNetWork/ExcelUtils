package lain.cn.common;

import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;

public class Common {
	public static final int BLANK_HEAD = 1;
	public static final int ERROR_ARGUMENT = 2;
	
	public static String titleTypeJudge(Cell cell) {
		switch (cell.getCellType()) {
		case BLANK:
			return "";
		case BOOLEAN:
			return cell.getBooleanCellValue()+"";
		case ERROR:
			return "";
		case FORMULA:
			return cell.getRichStringCellValue().toString();
		case NUMERIC:
			if(HSSFDateUtil.isCellDateFormatted(cell)) {				
				return new SimpleDateFormat("yyyy-MM-dd").format(cell.getDateCellValue());
			}else {
				return cell.getNumericCellValue()+"";
			}
		case STRING:
			return cell.getStringCellValue();
		case _NONE:
			return "";
		}
		return "";
	}
}
