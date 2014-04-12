package iExcelDemo;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

public class MySqlJDBCHelper implements IJDBCHelper {
	/***
	 * 获取连接
	 */
	public  Connection getConnection(LinkEntity entity) {
		Connection con = null;
		try {
			Class.forName(entity.getDriver());
			con = DriverManager.getConnection(entity.getUrl(), entity.getUser(), entity.getPassword());
			return con;
		} catch (Exception e) {

		}

		return null;
	}

	/***
	 * 获取结果集
	 */
	public  ResultSet getResultSet(Connection con, String sqlString) {
		Statement stmt = null;
		ResultSet rs = null;
		try {
			stmt = con.createStatement();
			rs = stmt.executeQuery(sqlString);
			
			return rs;
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	
		return null;
	}
}
