package iExcelDemo;

public class MainClass {

	public static void main(String[] args) {
		ExcelHelper helper = new ExcelHelper();
		String excelName = "test.xls";
		
		// ����excel
		helper.createExcel(excelName);
		// ��ȡexcel
		helper.readExcel(excelName);
	}
}
