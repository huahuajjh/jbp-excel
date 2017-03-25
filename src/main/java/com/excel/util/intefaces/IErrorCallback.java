package com.excel.util.intefaces;

import java.util.List;

import com.excel.util.model.ExcelRowError;

/**
 * 错误回掉
 * 
 * @author Bless
 * @time 2016/1/15 17:25
 * @version 1.0
 */
public interface IErrorCallback {
	
	void callback(List<ExcelRowError> errors);
}
