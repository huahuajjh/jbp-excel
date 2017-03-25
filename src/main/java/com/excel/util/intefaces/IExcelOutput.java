package com.excel.util.intefaces;

import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

import com.excel.util.model.ColAttrVal;
import com.excel.util.model.ColStyle;
import com.excel.util.model.ExcelColData;

/**
 * Excel写出操作对象
 * 
 * @author Bless
 * @version 1.0
 */
public interface IExcelOutput extends Closeable {
	/**
	 * 创建Sheet
	 * 
	 * @return Sheet的下标
	 */
	int createSheet();

	/**
	 * 创建Sheet
	 * 
	 * @param sheetName
	 *            Sheet的名称
	 * @return Sheet的下标
	 */
	int createSheet(String sheetName);

	/**
	 * 切换当前操作的Sheet
	 * 
	 * @param index
	 *            指定Sheet的下标 默认从0开始
	 * @return 是否切换成功（true为切换成功否则切换失败）
	 */
	boolean setSheetIndex(int index);
	
	/**
	 * 写入需要排除的列的数据
	 * @param indexs
	 */
	void setExcludeCol(Integer...indexs);
	
	/**
	 * 隐藏的列的数据
	 * @param indexs
	 */
	void hiddenCols(Integer...indexs);
	
	/**
	 * 删除指定的列
	 * @param indexs
	 */
	void removeCols(Integer...indexs);

	/**
	 * 获取当前正在操作的sheet的下表，-1表示当前没有Sheet被操作
	 * 
	 * @return Sheet的下标
	 */
	int getSheetIndex();

	/**
	 * 给当前操作的Sheet写入数据（默认从第0行开始）
	 * 
	 * @param type
	 *            指定数据的类型
	 * @param data
	 *            数据
	 * @throws IllegalAccessException
	 *             非法操作数据
	 * @throws NoSuchMethodException
	 *             没找到指定的方法
	 */
	<T> void writeData(Class<T> type, T data) throws IllegalAccessException, NoSuchMethodException;

	/**
	 * 给当前操作的Sheet写入数据
	 * 
	 * @param type
	 *            指定数据的类型
	 * @param data
	 *            数据
	 * @param rowIndex
	 *            行的下标（从0开始）
	 * @throws IllegalAccessException
	 *             非法操作数据
	 * @throws NoSuchMethodException
	 *             没找到指定的方法
	 */
	<T> void writeDatas(Class<T> type, List<T> data, int rowIndex) throws IllegalAccessException, NoSuchMethodException;

	/**
	 * 给当前操作的Sheet写入数据
	 * @param colDatas 数据集合
	 */
	void writeDatas(List<ExcelColData> colDatas);
	
	/**
	 * 给当前操作的Sheet写入数据
	 * @param type 指定数据的类型
	 * @param colAttrVals 属性对应的列
	 * @param rowIndex 所属的行
	 */
	<T> void writeDatas(Class<T> type,List<T> data,List<ColAttrVal> colAttrVals,int rowIndex) throws IllegalArgumentException, IllegalAccessException;
	
	/**
	 * 设置列的宽度
	 * 
	 * @param index
	 *            列的下标（从0开始）
	 * @param width
	 *            宽度
	 */
	void setColWidth(int index, int width);

	/**
	 * 设置行的高度
	 * 
	 * @param index
	 *            行的下标（从0开始）
	 * @param height
	 *            高度
	 */
	void setRowHeight(int index, Short heihgt);

	/**
	 * 导入样式
	 * 
	 * @param styles
	 *            样式数据
	 */
	void setStyleDatas(Map<String, ColStyle> styles);

	/**
	 * 导入央视
	 * 
	 * @param styleName
	 *            样式名称
	 * @param style
	 *            样式数据
	 */
	void setStyleData(String styleName, ColStyle style);

	/**
	 * 删除样式
	 * @param styleName 样式的名称
	 */
	void removeStyleAtName(String styleName);

	/**
	 * 删除所有样式
	 */
	void removeStyleAtAll();

	/**
	 * 把Excel写入指定的流
	 */
	void writeStream(OutputStream stream) throws IOException;
	
	/**
	 * 把Excel写入指定的流
	 */
	InputStream getInputStream() throws IOException;
	
	/**
	 * 给指定单元格写入数据
	 * @param rowIndex 行坐标（从0开始）
	 * @param colIndex 竖坐标（从0开始）
	 * @param val 写入的内容
	 */
	void setColVal(int rowIndex,int colIndex,Object val)  throws IllegalAccessException;
	
	/**
	 * 给指定单元格写入数据
	 * @param rowIndex 行坐标（从0开始）
	 * @param colIndex 竖坐标（从0开始）
	 * @param val 写入的内容
	 * @param rowMergeNum 横合并单元格的数量
	 * @param colMergeNum 竖合并单元格的数量
	 */
	void setColVal(int rowIndex,int colIndex,Object val,int rowMergeNum,int colMergeNum) throws IllegalAccessException;
	
}
