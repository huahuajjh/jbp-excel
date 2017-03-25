package com.excel.util;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.junit.Test;

import com.excel.util.annotation.OutputColAnnotation;

public class TestExport {
	
	@Test
	public void test1() throws IllegalAccessException, NoSuchMethodException, FileNotFoundException, IOException {
//		IExcelOutput output = ExcelFactory.getExcelOutput();
//		output.writeDatas(TestClas.class, Arrays.asList(new TestClas("A", 0), new TestClas("B", 1), new TestClas("C", 2)), 0);
//		
//		output.writeStream(new FileOutputStream("D:\\1.xlsx"));
		
	}
	
	
	class TestClas {
		@OutputColAnnotation(colCoord = 0)
		public String name;
		@OutputColAnnotation(colCoord = 1)
		public int age;
		
		public TestClas(String name, int age) {
			this.name = name;
			this.age = age;
		}
	}
}
