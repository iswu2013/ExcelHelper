package iExcelDemo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelHelper {

	/**
	 * ������excel����
	 * 
	 * @param excelName
	 *            excel����
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
			HSSFRow row = sheet.createRow((short) 0);
			// ������һ����Ԫ��
			HSSFCell cell = row.createCell((short) 0);
			// ���õ�Ԫ������Ϊ�ַ�����
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			// Ϊ�����ڵ�Ԫ����д�����ģ������ַ�����ΪENCODING_UTF_16
			cell.setCellValue("���Գɹ�");
			FileOutputStream outputStream = new FileOutputStream(excelName);
			workbook.write(outputStream);
			outputStream.flush();
			outputStream.close();

		} catch (Exception e) {

		}
	}

	/**
	 * �����ݿ��еĽ����������exel��
	 * 
	 * @param resultSet �����
	 * @param excelName excel����
	 * @param sheetName excel��sheet����
	 */
	public void resultSetToExcel(ResultSet rs, String excelName,
			String sheetName) {
		try {
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet sheet = workbook.createSheet();
			workbook.setSheetName(0, sheetName);
			HSSFRow row = sheet.createRow((short)0);
			HSSFCell cell;
			ResultSetMetaData md;
			md = rs.getMetaData();
			int colCount = md.getColumnCount();
			
			for(int i = 1;i <= colCount;i++) {
				cell = row.createCell((short)(i -1));
				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				cell.setCellValue(md.getColumnLabel(i));
			}
			
			int iRow = 1;
			while(rs.next()) {
				row = sheet.createRow((short)iRow);
				for(int j = 1; j <= colCount;j++) {
					cell = row.createCell((short)(j-1));
					cell.setCellType(HSSFCell.CELL_TYPE_STRING);
					cell.setCellValue(rs.getObject(j).toString());
				}
				iRow++;
			}
			
			FileOutputStream outputStream = new FileOutputStream(excelName);
			workbook.write(outputStream);
			outputStream.flush();
			outputStream.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * ��excel�ж�ȡ����
	 * 
	 * @param excelName
	 */
	public void readExcel(String excelName) {
		try {
			FileInputStream inputStream = new FileInputStream(excelName);
			HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
			HSSFSheet sheet = workbook.getSheet("firstSheet");
			HSSFRow row = sheet.getRow(0);
			HSSFCell cell = row.getCell((short) 0);
			System.out.println(cell.getStringCellValue());

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
