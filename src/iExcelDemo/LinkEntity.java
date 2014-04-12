package iExcelDemo;
/**
 * 链接实体对象，封装了链接元数据
 * 
 * @author wujinsong
 * 
 */
public class LinkEntity {
	// 用户名
	private String user;
	
	// 密码
	private String password;
	
	// url
	private String url;
	
	// 驱动
	private String driver;

	public String getUser() {
		return user;
	}

	public void setUser(String user) {
		this.user = user;
	}

	public String getPassword() {
		return password;
	}

	public void setPassword(String password) {
		this.password = password;
	}

	public String getUrl() {
		return url;
	}

	public void setUrl(String url) {
		this.url = url;
	}

	public String getDriver() {
		return driver;
	}

	public void setDriver(String driver) {
		this.driver = driver;
	}
}
