# ExcelUtils
根据自己的需求编写的一个Excel处理工具
可以进行类型判断，对各列进行类型投票，占多数的为列类型，并能筛选出类型异常的行输出。
比如一列数字中出现了一个字符串，该字符串所在行就不会出现在结果集里，可以通过.getErroData()输出。
使用方法如下：
	
	//输入流作为初始化参数，Excel版本03版，07版都可以
	ExcelHandler handler = new ExcelHandler(new FileInputStream("D:/Lain.xlsx"));	
	handler.getSheetNames();//可以获得所有工作簿名列表，内容为空的表会忽略
	SheetHandler sheetHandler = handler.getSheetHandlerByName("员工工资表");//根据工作簿名获取工作簿处理器
	sheetHandler.getErrorData();//获取错误行
	sheetHandler.getTypes();//获取列类型，分为三大类NUM，TEXT，DATE
	sheetHandler.getSheetData();//获取内容
支持错位的表格，比如表头位置不在(0,0)，而是在(22,33),依然可以正常读取数据。
以表头为列宽度的依据，表头不允许为空，表头首尾的空单元格会被忽略，而在中间出现的空单元格，将会抛出表头不能为空的异常。
