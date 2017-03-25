package com.excel.util.implemented;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.ss.util.CellRangeAddressBase;

import com.excel.util.annotation.InputColAnnotation;
import com.excel.util.error.ColErrorModel;
import com.excel.util.error.ConvertErrorException;
import com.excel.util.error.RowErrorModel;
import com.excel.util.intefaces.IErrorCallback;
import com.excel.util.intefaces.IExcelInput;
import com.excel.util.model.ExcelRowError;
import com.excel.util.model.RowReader;
import com.excel.util.tool.ExcelReaderUtil;

/**
 * POI封装
 * 
 * @author Bless
 * @version 1.0
 */
public class POIExcelInput implements IExcelInput {

	private static final Map<Class<?>, Field[]> classFields = new HashMap<Class<?>, Field[]>();
	private static final String[] cellVals = new String[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L",
			"M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

	private Map<String, Map<String, Boolean>> dataKeyMap;
	private IErrorCallback errorCallback;
	private List<Map<Integer, Map<Integer, String>>> excelDatas;
	private Map<Integer, Map<Integer, String>> sheet;
	private List<List<CellRangeAddressBase>> cellRanges;
	private List<CellRangeAddressBase> cellRange;
	private int errorNum = 0;

	public POIExcelInput(InputStream stream) throws Exception {
		RowReader rowRead = new RowReader();
		ExcelReaderUtil.readExcel(rowRead, stream);
		this.excelDatas = rowRead.getExcelDatas();
		this.cellRanges = rowRead.getCellRangeAddress();
		setSheetIndex(0);
	}

	public POIExcelInput(InputStream stream, int errorNum) throws Exception {
		this(stream);
		this.errorNum = errorNum;
	}

	public POIExcelInput(InputStream stream, IErrorCallback errorCallback) throws Exception {
		this(stream);
		this.errorCallback = errorCallback;
	}

	public POIExcelInput(InputStream stream, int errorNum, IErrorCallback errorCallback) throws Exception {
		this(stream, errorCallback);
		this.errorNum = errorNum;
	}

	@Override
	public Boolean setSheetIndex(int sheetIndex) {
		if (this.excelDatas.size() < sheetIndex || sheetIndex < 0) {
			return false;
		}
		this.sheet = this.excelDatas.get(sheetIndex);
		this.cellRange = this.cellRanges.get(sheetIndex);
		return true;
	}

	@Override
	public <T> T getData(Class<T> type) throws InstantiationException, ConvertErrorException, NoSuchMethodException {
		Map<Integer, Map<Integer, ColErrorModel>> errors = new HashMap<Integer, Map<Integer, ColErrorModel>>();
		this.dataKeyMap = new HashMap<String, Map<String, Boolean>>();
		try {
			return getData(type, 0);
		} catch (RowErrorModel e) {
			errors.put(e.getRowNum() + 1, e.getColErrors());
			// throw new ConvertErrorException(errors, "转换数据错误，你可通过错误的 getErrors
			// 方法获取错误详细");
		}
		if (errors.keySet().size() > 0 && errorCallback != null) {
			errorCallback.callback(conver(errors));
		}
		return null;
	}

	@Override
	public <T> List<T> getDatas(Class<T> type, int startRowIndex)
			throws InstantiationException, ConvertErrorException, NoSuchMethodException {
		Map<Integer, Map<Integer, ColErrorModel>> errors = new HashMap<Integer, Map<Integer, ColErrorModel>>();
		this.dataKeyMap = new HashMap<String, Map<String, Boolean>>();
		List<T> datas = new ArrayList<T>();
		if (this.sheet != null && startRowIndex > -1) {
			for (Entry<Integer, Map<Integer, String>> sheetData : this.sheet.entrySet()) {
				if (sheetData.getKey() >= startRowIndex) {
					try {
						T data = getData(type, sheetData.getKey());
						if (data != null) {
							datas.add(data);
						}
					} catch (RowErrorModel e) {
						errors.put(e.getRowNum() + 1, e.getColErrors());
					}
				}
				if (errorNum > 0 && errors.keySet().size() >= errorNum) {
					break;
				}
			}
		}
		if (errors.size() > 0) {
			if (errorCallback != null) {
				errorCallback.callback(conver(errors));
			}
			// throw new ConvertErrorException(errors, "转换数据错误，你可通过错误的 getErrors
			// 方法获取错误详细");
		}
		return datas;
	}

	@Override
	public <T> List<T> getDatas(Class<T> type, int startRowIndex, int number)
			throws InstantiationException, ConvertErrorException, NoSuchMethodException {
		Map<Integer, Map<Integer, ColErrorModel>> errors = new HashMap<Integer, Map<Integer, ColErrorModel>>();
		this.dataKeyMap = new HashMap<String, Map<String, Boolean>>();
		List<T> datas = new ArrayList<T>();
		if (this.sheet != null && number > 0 && startRowIndex > -1) {
			for (Entry<Integer, Map<Integer, String>> sheetData : this.sheet.entrySet()) {
				if (sheetData.getKey() >= startRowIndex) {
					try {
						T data = getData(type, sheetData.getKey());
						if (data != null) {
							datas.add(data);
						}
					} catch (RowErrorModel e) {
						errors.put(e.getRowNum() + 1, e.getColErrors());
					}
				}
				if (datas.size() >= number || (errorNum > 0 && errors.keySet().size() >= errorNum)) {
					break;
				}
			}
		}
		if (errors.size() > 0) {
			if (errorCallback != null) {
				errorCallback.callback(conver(errors));
			}
			// throw new ConvertErrorException(errors, "转换数据错误，你可通过错误的 getErrors
			// 方法获取错误详细");
		}
		return datas;
	}
	
	@Override
	public <T> List<T> getDatasExcludeAfterNum(Class<T> type, int startRowIndex, int rowNum) throws InstantiationException, NoSuchMethodException {
		Map<Integer, Map<Integer, ColErrorModel>> errors = new HashMap<Integer, Map<Integer, ColErrorModel>>();
		this.dataKeyMap = new HashMap<String, Map<String, Boolean>>();
		List<T> datas = new ArrayList<T>();
		if (this.sheet != null && rowNum > 0 && startRowIndex > -1) {
			Set<Integer> rowIndexSet = sheet.keySet();
			Integer max = 0;
			for (int i : rowIndexSet) {
				max = max>i?max:i;
			}
			max = max - rowNum;
			for (Entry<Integer, Map<Integer, String>> sheetData : this.sheet.entrySet()) {
				if (sheetData.getKey() >= startRowIndex) {
					try {
						T data = getData(type, sheetData.getKey());
						if (data != null) {
							datas.add(data);
						}
					} catch (RowErrorModel e) {
						errors.put(e.getRowNum() + 1, e.getColErrors());
					}
				}
				if (sheetData.getKey() >= max || (errorNum > 0 && errors.keySet().size() >= errorNum)) {
					break;
				}
			}
		}
		if (errors.size() > 0) {
			if (errorCallback != null) {
				errorCallback.callback(conver(errors));
			}
			// throw new ConvertErrorException(errors, "转换数据错误，你可通过错误的 getErrors
			// 方法获取错误详细");
		}
		return datas;
	}

	@Override
	public void close() throws IOException {
	}

	private List<ExcelRowError> conver(Map<Integer, Map<Integer, ColErrorModel>> errors) {
		List<ExcelRowError> errList = new ArrayList<ExcelRowError>();
		for (Map.Entry<Integer, Map<Integer, ColErrorModel>> row : errors.entrySet()) {
			for (Map.Entry<Integer, ColErrorModel> cel : row.getValue().entrySet()) {
				ExcelRowError excelRowError = new ExcelRowError();
				ColErrorModel model = cel.getValue();
				excelRowError.setCelIndex(cel.getKey());
				excelRowError.setColVal(model.getColVal());
				excelRowError.setConvertType(model.getConvertType().getName());
				excelRowError.setConvertValue(model.getConvertValue());
				excelRowError.setRowIndex(row.getKey());
				excelRowError.setMsg(model.getMsg());
				errList.add(excelRowError);
			}
		}
		return errList;
	}

	private <T> T getData(Class<T> cla, int rowIndex)
			throws InstantiationException, RowErrorModel, NoSuchMethodException {
		if (rowIndex < 0 || this.sheet == null || isEmptyRow(cla,rowIndex))
			return null;
		T data = null;
		if (!classFields.containsKey(cla))
			classFields.put(cla, cla.getDeclaredFields());
		Field[] fields = classFields.get(cla);
		RowErrorModel rowErrorModel = null;
		for (Field field : fields) {
			if (field.isAnnotationPresent(InputColAnnotation.class)) {
				field.setAccessible(true);
				InputColAnnotation ica = field.getAnnotation(InputColAnnotation.class);
				String errorMsg = ica.converErrorMsg();
				String strValue = null;
				int y = 0, x = 0;
				try {
					y = ica.colCoord();
					x = ica.rowCoord() >= 0 ? ica.rowCoord() : rowIndex;
					if(ica.isAbandonColspanData()){
						if(this.isColspan(y, x)){
							data = null;
							break;
						}
					}
					if(ica.isAbandonRowspanData()){
						if(this.isRowspan(y, x)){
							data = null;
							break;
						}
					}
					if(ica.required() && !this.sheet.containsKey(x)){
						errorMsg = ica.requiredErrorMsg();
						throw new Exception();
					} else if (!this.sheet.containsKey(x)) {
						continue;
					}
					Map<Integer, String> datas = this.sheet.get(x);
					if(ica.required() && !datas.containsKey(y)){
						errorMsg = ica.requiredErrorMsg();
						throw new Exception();
					}else if (!datas.containsKey(y)){
						continue;
					}
					strValue = datas.get(y);
					if ((strValue == null || strValue.equals("")) && ica.required()) {
						errorMsg = ica.requiredErrorMsg();
						throw new Exception();
					} else if (strValue == null || strValue.equals("")) {
						continue;
					}
					if (ica.only() && strValue!= null && !strValue.equals("")) {
						String name = field.getName();
						if (!this.dataKeyMap.containsKey(name)) {
							this.dataKeyMap.put(name, new HashMap<String, Boolean>());
						}
						Map<String, Boolean> valMap = this.dataKeyMap.get(name);
						if (valMap.containsKey(strValue)) {
							errorMsg = ica.onlyErrorMsg();
							throw new Exception();
						} else {
							valMap.put(strValue, true);
						}
					}
					if (data == null) {
						try {
							data = cla.newInstance();
						} catch (Exception e) {
							throw new InstantiationException(cla.getName() + "类必须是一个可new的，并且有一个无参的构造函数。");
						}
					}
					String methodName = ica.convertValMethod();
					if (methodName != null && !methodName.trim().equals("")) {
						methodName = methodName.trim();
						try {
							Method method = cla.getDeclaredMethod(methodName, String.class);
							method.setAccessible(true);
							method.invoke(data, strValue);
						} catch (Exception e) {
							throw new NoSuchMethodException("没找到 " + methodName + " 该方法");
						}
					} else {
						Object value = convertVal(field.getType(), strValue);
						field.set(data, value);
					}
				} catch (InstantiationException e) {
					throw e;
				} catch (NoSuchMethodException e) {
					throw e;
				} catch (Exception e) {
					if (rowErrorModel == null)
						rowErrorModel = new RowErrorModel(x);
					ColErrorModel colError = new ColErrorModel(field.getType(), strValue, getCellVal(y), errorMsg);
					rowErrorModel.setColErrors(y, colError);
				}
			}
		}
		if (rowErrorModel != null) {
			throw rowErrorModel;
		}
		return data;
	}

	private <T> boolean isEmptyRow(Class<T> cla, int rowIndex){
		if (!classFields.containsKey(cla))
			classFields.put(cla, cla.getDeclaredFields());
		Field[] fields = classFields.get(cla);
		List<Boolean> colState = new ArrayList<>();
		for (Field field : fields) {
			if (field.isAnnotationPresent(InputColAnnotation.class)) {
				field.setAccessible(true);
				InputColAnnotation ica = field.getAnnotation(InputColAnnotation.class);
				int y = 0, x = 0;
				y = ica.colCoord();
				x = ica.rowCoord() >= 0 ? ica.rowCoord() : rowIndex;
				if(!this.sheet.containsKey(x)) {
					colState.add(true);
					break;
				}
				Map<Integer, String> datas = this.sheet.get(x);
				if (!datas.containsKey(y)){
					colState.add(true);
					continue;
				}
				String strValue = datas.get(y);
				if(strValue == null || strValue.trim().equals("")){
					colState.add(true);
					continue;
				}
				colState.add(false);
			}
		}
		for (Boolean state : colState) {
			if(!state){
				return false;
			}
		}
		return true;
	}
	
	private boolean isColspan(int colIndex,int rowIndex){
		for (CellRangeAddressBase item : cellRange) {
			if(item.getFirstRow() <= rowIndex && item.getLastRow() >= rowIndex ){
				if(item.getFirstColumn() <= colIndex && item.getLastColumn() >= colIndex 
						&& item.getLastColumn() - item.getFirstColumn() > 0)
					return true;
			}
		}
		return false;
	}
	
	private boolean isRowspan(int colIndex,int rowIndex){
		for (CellRangeAddressBase item : cellRange) {
			if(item.getFirstRow() <= rowIndex && item.getLastRow() >= rowIndex)
				if(item.getFirstColumn() <= colIndex && item.getLastColumn() >= colIndex 
				&& item.getLastRow() - item.getFirstRow() > 0)
					return true;
		}
		return false;
	}
	
	private String getCellVal(int cellIndex) {
		cellIndex++;
		if(cellIndex <= 26){
			return cellVals[cellIndex - 1];
		} else {
			int indexB = cellIndex / 26;
			int indexA = cellIndex % 26;
			String val = cellVals[indexB - 1];
			val = val + cellVals[indexA - 1];
			return val;
		}
	}

	private Object convertVal(Class<?> cla, String val) throws ParseException {
		if (cla.equals(Double.class) || cla.equals(double.class)) {
			return Double.parseDouble(val);
		} else if (cla.equals(Float.class) || cla.equals(float.class)) {
			return Float.parseFloat(val);
		} else if (cla.equals(Integer.class) || cla.equals(int.class)) {
			int index = val.indexOf(".");
			String str = "";
			if (index > -1) {
				str = val.substring(0, index);
			} else {
				str = val;
			}
			return Integer.parseInt(str);
		} else if (cla.equals(Long.class) || cla.equals(long.class)) {
			int index = val.indexOf(".");
			String str = "";
			if (index > -1) {
				str = val.substring(0, index);
			} else {
				str = val;
			}
			return Long.parseLong(str);
		} else if (cla.equals(Short.class) || cla.equals(short.class)) {
			int index = val.indexOf(".");
			String str = "";
			if (index > -1) {
				str = val.substring(0, index);
			} else {
				str = val;
			}
			return Short.parseShort(str);
		} else if (cla.equals(Date.class)) {
			Double valD = Double.parseDouble(val);
			Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(valD);
			return date;
		} else if (cla.equals(String.class)) {
			return val;
		} else if (cla.equals(BigDecimal.class)) {
			return BigDecimal.valueOf(Double.parseDouble(val));
		} else {
			throw new ParseException("转换类型失败", 0);
		}
	}
}
