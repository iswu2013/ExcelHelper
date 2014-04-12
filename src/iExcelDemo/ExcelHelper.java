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
	 * ������excel����
	 * 
	 * @param excelName
	 */
	public void createExcel(String excelName) {
		try {
			// ��������������
			HSSFWorkbook workbook = new HSSFWorkbook();
			// �������������
			HSSFSheet sheet = workbook.createSheet();
			// ���õ�һ�������������ΪfirstShee
			// Ϊ�˹�������֧������
			workbook.setSheetName(0, "firstSheet");
			// ����һ��
			HSSFRow row = sheet.createRow((short)0);
			// ������һ����Ԫ��
			HSSFCell cell = row.createCell((short)0);
			// ���õ�Ԫ������Ϊ�ַ�����
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			// Ϊ�����ڵ�Ԫ����д�����ģ������ַ�����ΪENCODING_UTF_16
			cell.setCellValue("���Գɹ�");
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
