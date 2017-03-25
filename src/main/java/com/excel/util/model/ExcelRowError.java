package com.excel.util.model;

/**
 * 行列的错误数据
 * @author Bless
 * @time 2016/1/15 17:35
 * @version 1.0
 */
public class ExcelRowError {
	//行号
	private int rowIndex;
	//列号
	private int celIndex;
	// 转换的类型
	private String convertType;
	// 转换的数据
	private String convertValue;
	//列所处的下标
	private String colVal;
	//错误描述
	private String msg;
	
	public String getMsg() {
		return msg;
	}
	public void setMsg(String msg) {
		this.msg = msg;
	}
	public int getRowIndex() {
		return rowIndex;
	}
	public void setRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
	}
	public int getCelIndex() {
		return celIndex;
	}
	public void setCelIndex(int celIndex) {
		this.celIndex = celIndex;
	}
	public String getConvertType() {
		return convertType;
	}
	public void setConvertType(String convertType) {
		this.convertType = convertType;
	}
	public String getConvertValue() {
		return convertValue;
	}
	public void setConvertValue(String convertValue) {
		this.convertValue = convertValue;
	}
	public String getColVal() {
		return colVal;
	}
	public void setColVal(String colVal) {
		this.colVal = colVal;
	}
}
