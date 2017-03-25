package com.excel.util.error;

/**
 * 列错误数据模型
 * 
 * @author Bless
 * @version 1.0
 */
public final class ColErrorModel {
	// 转换的类型
	private Class<?> convertType;
	// 转换的数据
	private String convertValue;
	//列所处的下标
	private String colVal;
	//错误描述
	private String msg;
	
	/**
	 * 获取错误描述
	 * @return 错误描述
	 */
	public String getMsg() {
		return msg;
	}

	/**
	 * 获取转换错误的类型
	 * @return Class类型
	 */
	public Class<?> getConvertType() {
		return convertType;
	}

	/**
	 * 获取转换错误的数据
	 * @return 数据
	 */
	public String getConvertValue() {
		return convertValue;
	}

	/**
	 * 获取当前列所处的下标
	 * @return 列的下标
	 */
	public String getColVal() {
		return colVal;
	}
	
	/**
	 * 创建列错误对象
	 * @param 转换错误的类型
	 * @param 转换错误的内容
	 * @param 列所处下标
	 */
	public ColErrorModel(Class<?> convertType, String convertValue, String colVal,String msg) {
		super();
		this.msg = msg;
		this.convertType = convertType;
		this.convertValue = convertValue;
		this.colVal = colVal;
	}
}
