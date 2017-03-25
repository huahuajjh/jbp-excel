package com.excel.util.intefaces;

import java.util.List;
import java.util.Map;

import org.apache.poi.ss.util.CellRangeAddressBase;

/**
 * 
 * @author Bless
 * @version 1.0
 */
public interface IRowReader {
	
	/**业务逻辑实现方法
	 * @param sheetIndex
	 * @param curRow
	 * @param rowlist
	 */
	public void getRows(int sheetIndex,int curRow, Map<Integer, String> rowlist);
	
	/**
	 * 写入合并的数据
	 * @param cellRangeAddressBases
	 */
	public void setCellRangeAddress(int sheetIndex,List<CellRangeAddressBase> cellRangeAddressBases);
	
	/**
	 * 读取合并的数据
	 * @return
	 */
	public List<List<CellRangeAddressBase>> getCellRangeAddress();
}

