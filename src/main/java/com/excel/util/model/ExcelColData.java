package com.excel.util.model;

public class ExcelColData {
	/**单元格内容*/
	private String value;
	/**行合并的数量*/
	private int colspan;
	/**列合并的数量*/
	private int rowspan;
	/**行坐标*/
	private int x;
	/**列坐标*/
	private int y;

	public String getValue() {
		return value;
	}
	public void setValue(String value) {
		this.value = value;
	}
	public int getColspan() {
		return colspan;
	}
	public void setColspan(int colspan) {
		this.colspan = colspan;
	}
	public int getRowspan() {
		return rowspan;
	}
	public void setRowspan(int rowspan) {
		this.rowspan = rowspan;
	}
	public int getX() {
		return x;
	}
	public void setX(int x) {
		this.x = x;
	}
	public int getY() {
		return y;
	}
	public void setY(int y) {
		this.y = y;
	}
}
