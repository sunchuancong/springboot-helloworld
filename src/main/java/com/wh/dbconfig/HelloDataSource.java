package com.wh.dbconfig;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

@Component
@ConfigurationProperties("helloDataSource")
public class HelloDataSource {

	private String driver;
	private String url;
	private String userName;
	private String password;
	
	public String getUrl() {
		return url;
	}
	public void setUrl(String url) {
		this.url = url;
	}
	public String getUserName() {
		return userName;
	}
	public void setUserName(String userName) {
		this.userName = userName;
	}
	public String getPassword() {
		return password;
	}
	public void setPassword(String password) {
		this.password = password;
	}
	public String getDriver() {
		return driver;
	}
	public void setDriver(String driver) {
		this.driver = driver;
	}
	@Override
	public String toString() {
		return "HelloDataSource [driver=" + driver + ", url=" + url + ", userName=" + userName + ", password="
				+ password + "]";
	}
	
}
