package iExcelDemo;

public class MainClass {

	public static void main(String[] args) {
		ExcelHelper helper = new ExcelHelper();
		String excelName = "test.xls";
		
		// ²úÉúexcel
		helper.createExcel(excelName);
		// ¶ÁÈ¡excel
		helper.readExcel(excelName);
	}
}
