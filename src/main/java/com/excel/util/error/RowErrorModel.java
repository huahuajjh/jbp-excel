package com.excel.util.error;

import java.util.HashMap;
import java.util.Map;

/**
 * 行错误模型对象
 * 
 * @author Bless
 * @version 1.0
 */
public final class RowErrorModel extends Exception {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	// 列错误数据集合
	private final Map<Integer, ColErrorModel> colErrors;

	// 所处的行下标(从0开始)
	private int rowNum;

	/**
	 * 获取错误列详细
	 * 
	 * @return 错误列详细
	 */
	public Map<Integer, ColErrorModel> getColErrors() {
		return colErrors;
	}

	/**
	 * 写入列错误数据
	 * 
	 * @param colIndex
	 *            列所处的下标
	 * @param colError
	 *            错误详细
	 */
	public void setColErrors(Integer colIndex, ColErrorModel colError) {
		colErrors.put(colIndex, colError);
	}

	/**
	 * 获取当前行所处的下标（从0开始）
	 * 
	 * @return 行下标
	 */
	public int getRowNum() {
		return rowNum;
	}

	/**
	 * 创建行错误对象模型
	 * 
	 * @param rowNum 所处的行下标
	 */
	public RowErrorModel(int rowNum) {
		super();
		colErrors = new HashMap<Integer, ColErrorModel>();
		this.rowNum = rowNum;
	}
}
