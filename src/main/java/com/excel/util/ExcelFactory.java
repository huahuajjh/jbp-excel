package com.excel.util;

import java.io.IOException;
import java.io.InputStream;

import com.excel.util.implemented.POIExcelInput;
import com.excel.util.implemented.POIExcelOutput;
import com.excel.util.intefaces.IErrorCallback;
import com.excel.util.intefaces.IExcelInput;
import com.excel.util.intefaces.IExcelOutput;

/**
 * Excel操作对象工厂
 * 
 * @author Bless
 * @version 1.0
 */
public final class ExcelFactory {
	/**
	 * 获取Excel读取操作对象。
	 * 
	 * @author Bless
	 * @param inputStream
	 *            Excel文件流
	 * 
	 * @return Excel操作对象
	 * @throws IOException
	 *             传过来的文件流并非是Excel文件。
	 */
	public static IExcelInput getExcelInput(InputStream inputStream) throws IOException {
		try {
			return new POIExcelInput(inputStream);
		} catch (Exception e) {
			throw new IOException("该流并非是Excel文件。");
		}
	}
	
	/**
	 * 获取Excel读取操作对象。
	 * @param inputStream Excel文件流
	 * @param errNum 允许错误的条数，0以下为全部
	 * @return Excel操作对象
	 * @throws IOException
	 */
	public static IExcelInput getExcelInput(InputStream inputStream,int errNum) throws IOException{
		try {
			return new POIExcelInput(inputStream,errNum);
		} catch (Exception e) {
			throw new IOException("该流并非是Excel文件。");
		}
	}
	
	/**
	 * 获取Excel读取操作对象。
	 * @param inputStream Excel文件流
	 * @param errorCallback 错误回调
	 * @return
	 * @throws IOException
	 */
	public static IExcelInput getExcelInput(InputStream inputStream,IErrorCallback errorCallback) throws IOException{
		try {
			return new POIExcelInput(inputStream,errorCallback);
		} catch (Exception e) {
			throw new IOException("该流并非是Excel文件。");
		}
	}
	
	/**
	 * 获取Excel读取操作对象。
	 * @param inputStream Excel文件流
	 * @param errNum 允许错误的条数，0以下为全部
	 * @param errorCallback 错误回调
	 * @return Excel操作对象
	 * @throws IOException
	 */
	public static IExcelInput getExcelInput(InputStream inputStream,int errNum,IErrorCallback errorCallback) throws IOException{
		try {
			return new POIExcelInput(inputStream,errNum,errorCallback);
		} catch (Exception e) {
			throw new IOException("该流并非是Excel文件。");
		}
	}

	/**
	 * 获取Excel写入操作对象。
	 * 
	 * @author Bless
	 * @return Excel操作对象
	 */
	public static IExcelOutput getExcelOutput() {
		return new POIExcelOutput();
	}
	
	/**
	 * 获取Excel写入操作对象。
	 * @param inputStream 模板流
	 * @return Excel操作对象
	 * @throws IOException
	 */
	public static IExcelOutput getExcelOutput(InputStream inputStream) throws IOException{
		try {
			return new POIExcelOutput(inputStream);
		} catch (Exception e) {
			throw new IOException("该流并非是Excel文件。");
		}
	}
}
