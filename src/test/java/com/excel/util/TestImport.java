package com.excel.util;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

import org.junit.Test;

import com.excel.util.error.ConvertErrorException;
import com.excel.util.intefaces.IExcelInput;

public class TestImport {
	@Test
	public void test1() throws IllegalAccessException, NoSuchMethodException, FileNotFoundException, IOException,
			InstantiationException, ConvertErrorException {
		IExcelInput input = ExcelFactory.getExcelInput(new FileInputStream("D:\\1.xlsx"));
		List<User> list = input.getDatas(User.class, 0);

		list.forEach(d -> {
			System.out.println(d.toString());
		});
	}

}
