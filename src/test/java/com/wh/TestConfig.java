package com.wh;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import com.wh.dbconfig.HelloDataSource;

@RunWith(SpringRunner.class)
@SpringBootTest
public class TestConfig {

	@Autowired
	private HelloDataSource helloDataSource;
	
	@Test
	public void test1(){
		System.out.println(helloDataSource);
	}
	
	
}
