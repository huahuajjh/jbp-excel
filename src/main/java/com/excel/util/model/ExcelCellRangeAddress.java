package com.excel.util.model;

import org.apache.poi.ss.util.CellRangeAddressBase;

public class ExcelCellRangeAddress extends CellRangeAddressBase {
	
	public ExcelCellRangeAddress(int firstRow, int lastRow, int firstCol, int lastCol) {
		super(firstRow, lastRow, firstCol, lastCol);
	}

}
