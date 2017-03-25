package com.excel.util;

import com.excel.util.annotation.InputColAnnotation;

public class User {
	@InputColAnnotation(colCoord = 0)
	private String name;
	
	@InputColAnnotation(colCoord = 1)
	private int age;
	
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public int getAge() {
		return age;
	}
	public void setAge(int age) {
		this.age = age;
	}
	@Override
	public String toString() {
		return "User [name=" + name + ", age=" + age + "]";
	}
	
	

}
