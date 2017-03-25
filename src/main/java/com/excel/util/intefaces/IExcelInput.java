package com.excel.util.intefaces;

import java.io.Closeable;
import java.util.List;

import com.excel.util.error.ConvertErrorException;

/**
 * Excel读取操作对象
 * 
 * @author Bless
 * @version 1.0
 */
public interface IExcelInput extends Closeable {

	/**
	 * 切换Sheet，切换到指定的下标
	 * 
	 * @param sheetIndex
	 *            Sheet的下标（从0开始）
	 * @return 是否切换成功（true成功否则失败）
	 */
	Boolean setSheetIndex(int sheetIndex);

	/**
	 * 按照指定类型从Sheet中获取数据。（默认从第0行开始获取）
	 * 
	 * @param type
	 *            指定获取数据的类型
	 * @return 数据模型，如果没数据，返回值为null
	 * @throws InstantiationException
	 *             创建指定类型的对象的时候发生错误，必须有一个公开的无参构造函数。
	 * @throws ConvertErrorException
	 *             转换数据错误，可以通过 getError 获取详细错误信息
	 */
	<T> T getData(Class<T> type) throws InstantiationException, ConvertErrorException,NoSuchMethodException;

	/**
	 * 按照指定类型从Sheet中获取所有数据。
	 * 
	 * @param type
	 *            指定数据的类型
	 * @param startRowIndex
	 *            从第几行开始（从0开始）
	 * @return 数据集合
	 * @throws InstantiationException
	 *             创建指定类型的对象的时候发生错误，必须有一个公开的无参构造函数。
	 * @throws ConvertErrorException
	 *             转换数据错误，可以通过 getError 获取详细错误信息
	 */
	<T> List<T> getDatas(Class<T> type, int startRowIndex) throws InstantiationException, ConvertErrorException,NoSuchMethodException;

	/**
	 * 按照指定类型从Sheet中获取数据。
	 * 
	 * @param type
	 *            指定数据的类型
	 * @param startRowIndex
	 *            从第几行开始（从0开始）
	 * @param number
	 *            获取数据的数量
	 * @return 数据集合
	 * @throws InstantiationException
	 *             创建指定类型的对象的时候发生错误，必须有一个公开的无参构造函数。
	 * @throws ConvertErrorException
	 *             转换数据错误，可以通过 getError 获取详细错误信息
	 */
	<T> List<T> getDatas(Class<T> type, int startRowIndex, int number)
			throws InstantiationException, ConvertErrorException,NoSuchMethodException;
	
	/**
	 * 按照指定类型从Sheet中获取数据。（排除后面多少行数据）
	 * 
	 * @param type 指定数据的类型
	 * @param startRowIndex 从第几行开始（从0开始）
	 * @param rowNum 排除的行数量
	 * @return
	 */
	<T> List<T> getDatasExcludeAfterNum(Class<T> type,int startRowIndex ,int rowNum) throws InstantiationException, NoSuchMethodException ;

}
