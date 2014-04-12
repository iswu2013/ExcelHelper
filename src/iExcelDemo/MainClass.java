package iExcelDemo;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;

public class MainClass {

	public static void main(String[] args) {
		ExcelHelper helper = new ExcelHelper();
		String excelName = "test.xls";
		
		// 产生excel
		helper.createExcel(excelName);
		// 读取excel
		helper.readExcel(excelName);
		
		try {
			readFromMySql();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	/**
	 * 读取数据到
	 * @throws SQLException
	 */
	public static void readFromMySql() throws SQLException {
		LinkEntity entity = new LinkEntity();
		entity.setDriver("com.mysql.jdbc.Driver");
		entity.setPassword("wujinsong");
		entity.setUrl("jdbc:mysql://localhost:3306/mysql");
		entity.setUser("root");

		MySqlJDBCHelper mySqlJDBCHelper = new MySqlJDBCHelper();
		Connection con = mySqlJDBCHelper.getConnection(entity);

		String sqlString = "select * from help_topic";
		ResultSet rs = mySqlJDBCHelper.getResultSet(con, sqlString);
		ExcelHelper excelHelper = new ExcelHelper();
		excelHelper.resultSetToExcel(rs, "C:\\iExcelDemo\\mysql.xls", "help_topic");
		ResultSetMetaData rsmd = rs.getMetaData();
		int j = 0;
		j = rsmd.getColumnCount();
		for (int k = 0; k < j; k++) {
			System.out.print(rsmd.getCatalogName(k + 1));
			System.out.print("\t");
		}
		System.out.println();
		while (rs.next()) {
			for (int i = 0; i < j; i++) {
				System.out.print(rs.getString(i + 1));
				System.out.print("\t");
			}
			System.out.println();
		}
	}
}
