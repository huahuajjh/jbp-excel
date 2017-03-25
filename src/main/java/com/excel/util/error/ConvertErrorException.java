package com.excel.util.error;

import java.util.Map;

/**
 * 转换错误数据集合对象，可以通过对象的 getErrors 方法获取详细的错误信息
 * 
 * @author Bless
 * @version 1.0
 */
public final class ConvertErrorException extends Exception {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private final Map<Integer, Map<Integer, ColErrorModel>> errors;
	private String msg;

	/**
	 * 创建转换错误集合对象
	 * 
	 * @param errors
	 *            错误详细
	 * @param msg
	 *            错误描述
	 */
	public ConvertErrorException(Map<Integer, Map<Integer, ColErrorModel>> errors, String msg) {
		this.errors = errors;
		this.msg = msg;
	}

	/**
	 * 获取错误详细
	 * 
	 * @return 错误详细
	 */
	public Map<Integer, Map<Integer, ColErrorModel>> getErrors() {
		return errors;
	}

	/**
	 * 获取错误描述
	 * 
	 * @return 描述
	 */
	@Override
	public String getMessage() {
		return this.msg;
	}

}
