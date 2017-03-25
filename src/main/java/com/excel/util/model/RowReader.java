package com.excel.util.model;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.util.CellRangeAddressBase;

import com.excel.util.intefaces.IRowReader;

/**
 * 
 * @author Bless
 * @version 1.0
 */
public class RowReader implements IRowReader {

	private Map<Integer, List<CellRangeAddressBase>> cellRangeAddressBases;
	private Map<Integer, Map<Integer, Map<Integer, String>>> excelDatas;

	public List<Map<Integer, Map<Integer, String>>> getExcelDatas() {
		List<Map<Integer, Map<Integer, String>>> sheets = new ArrayList<>();
		Set<Integer> keySet = excelDatas.keySet();
		Integer[] keys = keySet.toArray(new Integer[keySet.size()]);
		Arrays.sort(keys);
		for (Integer key : keys) {
			Map<Integer, Map<Integer, String>> mapData = excelDatas.get(key);
			if (mapData != null) {
				sheets.add(mapData);

				List<CellRangeAddressBase> list = cellRangeAddressBases.get(key);
				for (CellRangeAddressBase item : list) {
					int firstRow = item.getFirstRow();
					int lastRow = item.getLastRow();
					int filrstCol = item.getFirstColumn();
					int lastCol = item.getLastColumn();
					String val = null;
					if (mapData.containsKey(firstRow)) {
						Map<Integer, String> row = mapData.get(firstRow);
						if (row.containsKey(filrstCol)) {
							val = row.get(filrstCol);
						}
					}
					for (int j = firstRow; j <= lastRow; j++) {
						if (!mapData.containsKey(j)) {
							mapData.put(j, new HashMap<Integer, String>());
						}
						Map<Integer, String> row = mapData.get(j);
						for (int j2 = filrstCol; j2 <= lastCol; j2++) {
							row.put(j2, val);
						}
					}
				}
			}
		}
		return sheets;
	}

	public RowReader() {
		this.cellRangeAddressBases = new HashMap<>();
		this.excelDatas = new HashMap<Integer, Map<Integer, Map<Integer, String>>>();
	}

	/*
	 * 业务逻辑实现方法
	 * 
	 * @see com.eprosun.util.excel.IRowReader#getRows(int, int, java.util.List)
	 */
	public void getRows(int sheetIndex, int curRow, Map<Integer, String> rowlist) {
		if (!this.excelDatas.containsKey(sheetIndex)) {
			this.excelDatas.put(sheetIndex, new HashMap<Integer, Map<Integer, String>>());
		}
		Map<Integer, Map<Integer, String>> sheetData = this.excelDatas.get(sheetIndex);
		sheetData.put(curRow, rowlist);
	}

	@Override
	public void setCellRangeAddress(int sheetIndex, List<CellRangeAddressBase> cellRangeAddressBases) {
		this.cellRangeAddressBases.put(sheetIndex, cellRangeAddressBases);
	}

	@Override
	public List<List<CellRangeAddressBase>> getCellRangeAddress() {
		List<List<CellRangeAddressBase>> cellRanges = new ArrayList<>();
		Set<Integer> keySet = excelDatas.keySet();
		Integer[] keys = keySet.toArray(new Integer[keySet.size()]);
		Arrays.sort(keys);
		for (Integer key : keys) {
			cellRanges.add(cellRangeAddressBases.get(key));
		}
		return cellRanges;
	}
}
