package com.excel.util.implemented;

import java.io.ByteArrayInputStream;
import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.excel.util.annotation.OutputColAnnotation;
import com.excel.util.intefaces.IExcelOutput;
import com.excel.util.model.ColAttrVal;
import com.excel.util.model.ColFont;
import com.excel.util.model.ColStyle;
import com.excel.util.model.ExcelColData;
import com.excel.util.tool.SheetUtility;

/**
 * POI封装
 * 
 * @author Bless
 * @version 1.0
 */
public class POIExcelOutput implements IExcelOutput, Closeable {

	private static final Map<Class<?>, Field[]> classFields = new HashMap<Class<?>, Field[]>();

	private XSSFWorkbook xWorkbook;
	private InputStream inputStream;
	private Workbook workbook;
	private Sheet sheet;
	private String[] notRemoveStyleNames = new String[] { "defaultDate" };
	private Map<String, CellStyle> cellStyles = new HashMap<String, CellStyle>();
	private Set<Integer> excludeColIndexs = new HashSet<>();

	public POIExcelOutput() {
		this.workbook = new SXSSFWorkbook(10);
		init();
	}

	public POIExcelOutput(InputStream inputStream) throws IOException {
		ByteArrayOutputStream output = new ByteArrayOutputStream();
		output.write(inputStream);
		byte[] bytes = output.toByteArray();
		output.close();
		this.inputStream = new ByteArrayInputStream(bytes);
		this.xWorkbook = new XSSFWorkbook(this.inputStream);
		this.workbook = new SXSSFWorkbook(this.xWorkbook,10);
		init();
	}

	private void init() {
		DataFormat df = workbook.createDataFormat();
		CellStyle cellDateStyle = workbook.createCellStyle();
		cellDateStyle.setDataFormat(df.getFormat("m/d/yy"));
		cellStyles.put("defaultDate", cellDateStyle);
		CellStyle borderStyle = workbook.createCellStyle();
		borderStyle.setAlignment(org.apache.poi.ss.usermodel.CellStyle.ALIGN_CENTER);
		borderStyle.setBorderBottom(org.apache.poi.ss.usermodel.CellStyle.BORDER_THIN); //下边框    
		borderStyle.setBorderLeft(org.apache.poi.ss.usermodel.CellStyle.BORDER_THIN);//左边框    
		borderStyle.setBorderTop(org.apache.poi.ss.usermodel.CellStyle.BORDER_THIN);//上边框    
		borderStyle.setBorderRight(org.apache.poi.ss.usermodel.CellStyle.BORDER_THIN);//右边框 
		Font font = this.workbook.createFont();
		font.setFontHeight((short) 250);
		borderStyle.setFont(font);
		cellStyles.put("defaultBorder", borderStyle);
	}

	@Override
	public void close() throws IOException {
		if (this.inputStream != null) {
			this.inputStream.close();
		}
		this.workbook.close();
	}

	@Override
	public void setExcludeCol(Integer... indexs) {
		excludeColIndexs.clear();
		if (indexs != null) {
			for (Integer i : indexs) {
				excludeColIndexs.add(i);
			}
		}
	}

	@Override
	public void hiddenCols(Integer... indexs) {
		if (indexs == null)
			return;
		for (Integer i : indexs) {
			this.setColWidth(i, 0);
		}
	}

	@Override
	public int createSheet() {
		Sheet tempSheet = this.workbook.createSheet();
		return this.workbook.getSheetIndex(tempSheet);
	}

	@Override
	public int createSheet(String sheetName) {
		Sheet tempSheet = this.workbook.createSheet("sheetName");
		return this.workbook.getSheetIndex(tempSheet);
	}

	@Override
	public boolean setSheetIndex(int index) {
		int sheetNum = this.workbook.getNumberOfSheets();
		if (index >= sheetNum || index < 0)
			return false;
		this.sheet = this.workbook.getSheetAt(index);
		return true;
	}

	@Override
	public int getSheetIndex() {
		if (this.sheet == null)
			return -1;
		return workbook.getSheetIndex(this.sheet);
	}

	@Override
	public void setColVal(int rowIndex, int colIndex, Object val) throws IllegalAccessException {
		setColVal(rowIndex, colIndex, val, 1, 1);
	}

	@Override
	public void setColVal(int rowIndex, int colIndex, Object val, int rowMergeNum, int colMergeNum)
			throws IllegalAccessException {
		if (rowIndex < 0)
			rowIndex = 0;
		if (colIndex < 0)
			colIndex = 0;
		if (rowMergeNum < 0)
			rowMergeNum = 0;
		if (colMergeNum < 0)
			colMergeNum = 0;
		if (val == null)
			return;
		this.ifSheetIsZeroCreateSheet();
		Row row = this.sheet.getRow(rowIndex);
		if (row == null)
			row = this.sheet.createRow(rowIndex);
		Cell cell = row.getCell(colIndex);
		if (cell == null){
			cell = row.createCell(colIndex);
		}
		Class<? extends Object> valType = val.getClass();
		if (valType.equals(String.class) && val.toString().startsWith("=")) {
			cell.setCellFormula(val.toString().substring(1));
		} else {
			try {
				if (val != null && (valType == int.class || valType == Integer.class || valType == double.class
						|| valType == Double.class || valType == float.class || valType == Float.class
						|| valType == short.class || valType == Short.class || valType == Long.class
						|| valType == long.class || valType == BigDecimal.class)) {
					cell.setCellValue(Double.parseDouble(val.toString()));
				} else if (val != null && valType == Date.class) {
					cell.setCellStyle(this.cellStyles.get("defaultDate"));
					cell.setCellValue((Date) val);
				} else if (val != null) {
					cell.setCellValue(val.toString());
				}
			} catch (Exception e) {
				throw new IllegalAccessException("写入数据的时候出现错了。");
			}
		}
		// 合并单元格
		int colspan = colMergeNum > 1 ? colMergeNum - 1 : 0;
		int rowspan = rowMergeNum > 1 ? rowMergeNum - 1 : 0;
		if (colspan > 0 || rowspan > 0) {
			colspan = colIndex + colspan;
			rowspan = rowIndex + rowspan;
			CellRangeAddress cra = new CellRangeAddress(rowIndex, rowspan, colIndex, colspan);
			this.sheet.addMergedRegion(cra);
		}
	}

	@Override
	public <T> void writeData(Class<T> type, T data) throws IllegalAccessException, NoSuchMethodException {
		this.ifSheetIsZeroCreateSheet();
		this.writeData(0, type, data);
	}

	@Override
	public <T> void writeDatas(Class<T> type, List<T> data, int rowIndex)
			throws IllegalAccessException, NoSuchMethodException {
		this.ifSheetIsZeroCreateSheet();
		if (rowIndex < 0) {
			rowIndex = 0;
		}
		for (T t : data) {
			this.writeData(rowIndex, type, t);
			rowIndex++;
		}
	}

	@Override
	public void writeDatas(List<ExcelColData> colDatas) {
		this.ifSheetIsZeroCreateSheet();
		if (colDatas == null || colDatas.size() == 0)
			return;
		for (ExcelColData excelColData : colDatas) {
			int x = excelColData.getX();
			Row row = this.sheet.getRow(x);
			if (row == null) {
				row = this.sheet.createRow(x);
			}
			int y = excelColData.getY();
			Cell cell = row.getCell(y);
			if (cell == null) {
				cell = row.createCell(y);
			}
			cell.setCellValue(excelColData.getValue());
			int colspan = excelColData.getColspan() < 0 ? 0 : excelColData.getColspan();
			int rowspan = excelColData.getRowspan() < 0 ? 0 : excelColData.getRowspan();
			if (colspan > 0 || rowspan > 0) {
				colspan = y + colspan;
				rowspan = x + rowspan;
				CellRangeAddress cra = new CellRangeAddress(x, rowspan, y, colspan);
				this.sheet.addMergedRegion(cra);
				SheetUtility.setRegionStyle(this.cellStyles.get("defaultBorder"),cra,this.sheet);
			} else{
				cell.setCellStyle(this.cellStyles.get("defaultBorder"));
			}
		}
	}
	
	@Override
	public <T> void writeDatas(Class<T> type,List<T> data,List<ColAttrVal> colAttrVals,int rowIndex) throws IllegalArgumentException, IllegalAccessException {
		this.ifSheetIsZeroCreateSheet();
		if(data == null || data.size() == 0) return;
		if (rowIndex < 0) {
			rowIndex = 0;
		}
		if (!classFields.containsKey(type))
			classFields.put(type, type.getDeclaredFields());
		Field[] fields = classFields.get(type);
		Map<String, Field> nameFields = new HashMap<>();
		for (Field field : fields) {
			field.setAccessible(true);
			nameFields.put(field.getName(), field);
		}
		for (T t : data) {
			Row row = this.sheet.getRow(rowIndex);
			if(row == null) row = this.sheet.createRow(rowIndex);
			if (t == null)
				continue;
			for (ColAttrVal colAttrVal : colAttrVals) {
				Cell cell = row.getCell(colAttrVal.getColIndex());
				if(cell == null) cell = row.createCell(colAttrVal.getColIndex());
				String name = colAttrVal.getAttrName();
				if(!nameFields.containsKey(name)) continue;
				
				Field field = nameFields.get(name);
				Object val = field.get(t);
				Class<?> valType = field.getType();
				try {
					if(valType.equals(String.class) && val != null && val.toString().startsWith("=")){
						cell.setCellFormula(val.toString().substring(1));
					} else if (val != null && (valType == int.class || valType == Integer.class
							|| valType == double.class || valType == Double.class || valType == float.class
							|| valType == Float.class || valType == short.class || valType == Short.class
							|| valType == Long.class || valType == long.class || valType == BigDecimal.class)) {
						cell.setCellValue(Double.parseDouble(val.toString()));
					} else if (val != null && valType == Date.class) {
						cell.setCellStyle(this.cellStyles.get("defaultDate"));
						cell.setCellValue((Date) val);
					} else if (val != null) {
						cell.setCellValue(val.toString());
					}
				} catch (Exception e) {
					throw new IllegalAccessException("写入数据的时候出现错了。");
				}
			}
			rowIndex++;
		}
		sheet.setForceFormulaRecalculation(true);
	}

	@Override
	public void setColWidth(int index, int width) {
		if (index < 0)
			return;
		this.ifSheetIsZeroCreateSheet();
		this.sheet.setColumnWidth(index, width);
	}

	@Override
	public void setRowHeight(int index, Short heihgt) {
		if (index < 0)
			return;
		Row row = sheet.getRow(index);
		if (row == null) {
			row = sheet.createRow(index);
			return;
		}
		row.setHeight(heihgt);
	}

	@Override
	public void setStyleDatas(Map<String, ColStyle> styles) {
		if (styles == null) {
			return;
		}
		for (Map.Entry<String, ColStyle> style : styles.entrySet()) {
			this.setStyleData(style.getKey(), style.getValue());
		}
	}

	@Override
	public void setStyleData(String styleName, ColStyle style) {
		if (style == null) {
			return;
		}
		boolean state = true;
		for (String name : notRemoveStyleNames) {
			if (name == styleName) {
				state = false;
			}
		}
		if (state) {
			CellStyle cellStyle = this.workbook.createCellStyle();
			this.convertColStyleAtCellStyle(style, cellStyle);
			this.cellStyles.put(styleName, cellStyle);
		}
	}

	@Override
	public void removeStyleAtName(String styleName) {
		boolean state = true;
		for (String name : notRemoveStyleNames) {
			if (name == styleName) {
				state = false;
			}
		}
		if (state) {
			this.cellStyles.remove(styleName);
		}
	}

	@Override
	public void removeStyleAtAll() {
		Set<String> styleNames = this.cellStyles.keySet();
		for (String styleName : styleNames) {
			boolean state = true;
			for (String name : notRemoveStyleNames) {
				if (name == styleName) {
					state = false;
				}
			}
			if (state) {
				this.cellStyles.remove(styleName);
			}
		}
	}

	@Override
	public void writeStream(OutputStream stream) throws IOException {
		workbook.write(stream);
	}

	@Override
	public InputStream getInputStream() throws IOException {
		ByteArrayOutputStream output = new ByteArrayOutputStream();
		writeStream(output);
		byte[] bytes = output.toByteArray();
		output.close();
		return new ByteArrayInputStream(bytes);
	}

	@Override
	public void removeCols(Integer... indexs) {
		if (indexs == null || indexs.length == 0)
			return;
		for (int i = 0; i <= this.sheet.getLastRowNum(); i++) {
			Row row = this.sheet.getRow(i);
			for (Integer celIndex : indexs) {
				// SheetUtility.deleteColumn(sheet,celIndex);
				Cell cell = row.getCell(celIndex);
				CellStyle style = cell.getCellStyle();
				style.setHidden(true);
				style.setLocked(true);
			}
		}

	}

	private void ifSheetIsZeroCreateSheet() {
		int sheetNum = this.workbook.getNumberOfSheets();
		if (sheetNum <= 0) {
			this.createSheet();
		}
		if (this.sheet == null) {
			this.setSheetIndex(0);
		}
	}

	private <T> void writeData(int rowIndex, Class<T> cla, T data)
			throws IllegalAccessException, NoSuchMethodException {
		if (data == null)
			return;
		if (!classFields.containsKey(cla))
			classFields.put(cla, cla.getDeclaredFields());
		Field[] fields = classFields.get(cla);
		for (Field field : fields) {
			if (field.isAnnotationPresent(OutputColAnnotation.class)) {
				field.setAccessible(true);
				OutputColAnnotation oca = field.getAnnotation(OutputColAnnotation.class);
				// 获取row和cell对象
				int x = oca.rowCoord() > -1 ? oca.rowCoord() : rowIndex;
				Row row = this.sheet.getRow(x);
				if (row == null)
					row = this.sheet.createRow(x);
				int y = oca.colCoord() > -1 ? oca.colCoord() : 0;
				Cell cell = row.getCell(y);
				if (cell == null)
					cell = row.createCell(y);
				if (oca.formula() != null && !oca.formula().equals("")) {
					cell.setCellFormula(oca.formula());
				}
				// 写入数据
				Object val = field.get(data);
				Class<?> valType = field.getType();
				String methodName = oca.convertValMethod();
				if (methodName != null && !methodName.trim().equals("")) {
					methodName = methodName.trim();
					try {
						Method method = cla.getDeclaredMethod(methodName);
						val = method.invoke(data);
						valType = val.getClass();
					} catch (Exception e) {
						throw new NoSuchMethodException("没找到该方法");
					}
				}
				try {
					if (excludeColIndexs.contains(y)) {
						// 不需要对该行写入数据
					} else if (val != null && (valType == int.class || valType == Integer.class
							|| valType == double.class || valType == Double.class || valType == float.class
							|| valType == Float.class || valType == short.class || valType == Short.class
							|| valType == Long.class || valType == long.class || valType == BigDecimal.class)) {
						cell.setCellValue(Double.parseDouble(val.toString()));
					} else if (val != null && valType == Date.class) {
						cell.setCellStyle(this.cellStyles.get("defaultDate"));
						cell.setCellValue((Date) val);
					} else if (val != null) {
						cell.setCellValue(val.toString());
					}
				} catch (Exception e) {
					throw new IllegalAccessException("写入数据的时候出现错了。");
				}

				// 设置样式
				String styleDataName = oca.styleDataName();
				if (styleDataName != null && !styleDataName.trim().equals("")
						&& this.cellStyles.containsKey(styleDataName)) {
					cell.setCellStyle(this.cellStyles.get(styleDataName));
				}

				// 合并单元格
				int colspan = oca.colspan() > 1 ? oca.colspan() - 1 : 0;
				int rowspan = oca.rowspan() > 1 ? oca.rowspan() - 1 : 0;
				if (colspan > 0 || rowspan > 0) {
					colspan = y + colspan;
					rowspan = x + rowspan;
					CellRangeAddress cra = new CellRangeAddress(x, rowspan, y, colspan);
					this.sheet.addMergedRegion(cra);
				}
			}
		}
		sheet.setForceFormulaRecalculation(true);
	}

	private void convertColStyleAtCellStyle(ColStyle colStyle, CellStyle cellDateStyle) {
		cellDateStyle.setAlignment(colStyle.getAlignment());
		cellDateStyle.setBorderBottom(colStyle.getBorderBottom());
		cellDateStyle.setBorderLeft(colStyle.getBorderLeft());
		cellDateStyle.setBorderRight(colStyle.getBorderRight());
		cellDateStyle.setBorderTop(colStyle.getBorderTop());
		cellDateStyle.setBottomBorderColor(colStyle.getBottomBorderColor());
		if (colStyle.getDataFormat() != null && !colStyle.getDataFormat().trim().equals("")) {
			DataFormat df = workbook.createDataFormat();
			cellDateStyle.setDataFormat(df.getFormat(colStyle.getDataFormat()));
		}
		cellDateStyle.setFillBackgroundColor(colStyle.getFillBackgroundColor());
		cellDateStyle.setFillForegroundColor(colStyle.getFillForegroundColor());
		cellDateStyle.setFillPattern(colStyle.getFillPattern());
		if (colStyle.getFont() != null) {
			Font font = this.workbook.createFont();
			this.convertColFontAtCellFont(colStyle.getFont(), font);
			cellDateStyle.setFont(font);
		}
		cellDateStyle.setHidden(colStyle.getHidden());
		cellDateStyle.setIndention(colStyle.getIndention());
		cellDateStyle.setLeftBorderColor(colStyle.getLeftBorderColor());
		cellDateStyle.setLocked(colStyle.getLocked());
		cellDateStyle.setRightBorderColor(colStyle.getRightBorderColor());
		cellDateStyle.setRotation(colStyle.getRotation());
		cellDateStyle.setShrinkToFit(colStyle.getShrinkToFit());
		cellDateStyle.setTopBorderColor(colStyle.getTopBorderColor());
		cellDateStyle.setVerticalAlignment(colStyle.getVerticalAlignment());
		cellDateStyle.setWrapText(colStyle.getWrapText());
	}

	private void convertColFontAtCellFont(ColFont colFont, Font cellFont) {
		cellFont.setBold(colFont.getBold());
		cellFont.setBoldweight(colFont.getBoldweight());
		cellFont.setCharSet(colFont.getCharSet());
		cellFont.setColor(colFont.getColor());
		cellFont.setFontHeight(colFont.getFontHeight());
		cellFont.setFontHeightInPoints(colFont.getFontHeightInPoints());
		cellFont.setFontName(colFont.getFontName());
		cellFont.setItalic(colFont.getItalic());
		cellFont.setStrikeout(colFont.getStrikeout());
		cellFont.setTypeOffset(colFont.getTypeOffset());
		cellFont.setUnderline(colFont.getUnderline());
	}
}
