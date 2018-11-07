package lain.cn.excel.exception;

public class ExcelLoadException extends RuntimeException{
	private static final long serialVersionUID = 4509639862218558146L;
	
	public ExcelLoadException(String message) {
		super(message);
	}
	public ExcelLoadException() {
		super("读取Excel文件失败!");
	}
	
	
	
}
