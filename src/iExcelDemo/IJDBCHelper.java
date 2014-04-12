package iExcelDemo;
import java.sql.Connection;
import java.sql.ResultSet;

public interface IJDBCHelper {
	/**
	 * 获取连接
	 * 
	 * @return
	 */
	public abstract Connection getConnection(LinkEntity entity);

	/**
	 * 获取结果集
	 * @return
	 */
	public abstract ResultSet getResultSet(Connection connection,String sqlString);
}
