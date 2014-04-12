package iExcelDemo;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelHelper {

	/**
	 * 产生的excel名称
	 * 
	 * @param excelName
	 */
	public void createExcel(String excelName) {
		try {
			// 产生工作簿对象
			HSSFWorkbook workbook = new HSSFWorkbook();
			// 产生工作表对象
			HSSFSheet sheet = workbook.createSheet();
			// 设置第一个工作表的名称为firstShee
			// 为了工作表能支持中文
			workbook.setSheetName(0, "firstSheet");
			// 产生一行
			HSSFRow row = sheet.createRow((short)0);
			// 产生第一个单元格
			HSSFCell cell = row.createCell((short)0);
			// 设置单元格内容为字符串型
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			// 为了能在单元格中写入中文，设置字符编码为ENCODING_UTF_16
			cell.setCellValue("测试成功");
			FileOutputStream outputStream = new FileOutputStream(excelName);
			workbook.write(outputStream);
			outputStream.flush();
			outputStream.close();
			
		} catch(Exception e)
		{
			
		}
	}
	
	public void readExcel(String excelName) {
		try {
			FileInputStream inputStream = new FileInputStream(excelName);
			HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
			HSSFSheet sheet = workbook.getSheet("firstSheet");
			HSSFRow row = sheet.getRow(0);
			HSSFCell cell = row.getCell((short)0);
			System.out.println(cell.getStringCellValue());
			
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
