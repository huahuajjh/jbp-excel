package com.excel.util.tool;

import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.ws.WebServiceException;

import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder.SheetRecordCollectingListener;
import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BlankRecord;
import org.apache.poi.hssf.record.BoolErrRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.FormulaRecord;
import org.apache.poi.hssf.record.LabelRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.hssf.record.StringRecord;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddressBase;

import com.excel.util.intefaces.IRowReader;

/**
 * 抽象Excel2003读取器，通过实现HSSFListener监听器，采用事件驱动模式解析excel2003
 * 中的内容，遇到特定事件才会触发，大大减少了内存的使用。
 *
 */

/**
 * 抽象Excel2003读取器，通过实现HSSFListener监听器，采用事件驱动模式解析excel2003
 * 中的内容，遇到特定事件才会触发，大大减少了内存的使用。
 *
 */
/**
 * 抽象Excel2003读取器，通过实现HSSFListener监听器，采用事件驱动模式解析excel2003
 * 中的内容，遇到特定事件才会触发，大大减少了内存的使用。
 *
 */
/**
 * 
 * @author Bless
 * @version 1.0
 */
@SuppressWarnings("restriction")
public class Excel2003Reader implements HSSFListener,Closeable {

	private POIFSFileSystem fs;

	HSSFWorkbook workbook;
	
	/** Should we output the formula, or the value it has? */
	private boolean outputFormulaValues = true;

	/** For parsing Formulas */
	private SheetRecordCollectingListener workbookBuildingListener;
	// excel2003工作薄
	private HSSFWorkbook stubWorkbook;

	// Records we pick up as we process
	private SSTRecord sstRecord;
	private FormatTrackingHSSFListener formatListener;

	// 表索引
	private int sheetIndex = -1;
	private BoundSheetRecord[] orderedBSRs;
	private ArrayList boundSheetRecords = new ArrayList();

	// For handling formulas with string results
	private int nextRow;
	private int nextColumn;
	private boolean outputNextStringRecord;
	// 当前行
	private int curRow = 0;
	// 存储行记录的容器
	// private List<String> rowlist = new ArrayList<String>();
	private Map<Integer, String> rowlist = new HashMap<Integer, String>();

	@SuppressWarnings("unused")
	private String sheetName;

	private IRowReader rowReader;

	public void setRowReader(IRowReader rowReader) {
		this.rowReader = rowReader;
	}

	/**
	 * 遍历excel下所有的sheet
	 * 
	 * @throws IOException
	 */
	public void process(InputStream inputStream) throws IOException {
		this.fs = new POIFSFileSystem(inputStream);
		workbook = new HSSFWorkbook(fs);
//		if (this.rowReader != null) {
//			int sheetSize = workbook.getNumberOfSheets();
//			List<List<CellRangeAddressBase>> cellRangeAddressBases = new ArrayList<List<CellRangeAddressBase>>();
//			for (int i = 0; i < sheetSize; i++) {
//				List<CellRangeAddressBase> temps = new ArrayList<>();
//				Sheet sheet = workbook.getSheetAt(i);
//				int sheetMergeCount = sheet.getNumMergedRegions();
//				for (int j = 0; j < sheetMergeCount; j++) {
//					temps.add(sheet.getMergedRegion(j));
//				}
//				cellRangeAddressBases.add(temps);
//			}
//			this.rowReader.setCellRangeAddress(cellRangeAddressBases);
//		}
		MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
		formatListener = new FormatTrackingHSSFListener(listener);
		HSSFEventFactory factory = new HSSFEventFactory();
		HSSFRequest request = new HSSFRequest();
		if (outputFormulaValues) {
			request.addListenerForAllRecords(formatListener);
		} else {
			workbookBuildingListener = new SheetRecordCollectingListener(formatListener);
			request.addListenerForAllRecords(workbookBuildingListener);
		}
		factory.processWorkbookEvents(request, fs);
	}

	/**
	 * HSSFListener 监听方法，处理 Record
	 */
	@SuppressWarnings("unchecked")
	public void processRecord(Record record) {
		int thisRow = -1;
		int thisColumn = -1;
		String thisStr = null;
		String value = null;
		switch (record.getSid()) {
		case BoundSheetRecord.sid:
			boundSheetRecords.add(record);
			break;
		case BOFRecord.sid:
			BOFRecord br = (BOFRecord) record;
			if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
				// 如果有需要，则建立子工作薄
				if (workbookBuildingListener != null && stubWorkbook == null) {
					stubWorkbook = workbookBuildingListener.getStubHSSFWorkbook();
				}
				sheetIndex++;
				Sheet sheet = workbook.getSheetAt(sheetIndex);
				
				int sheetMergeCount = sheet.getNumMergedRegions();
				List<CellRangeAddressBase> temps = new ArrayList<>();
				for (int j = 0; j < sheetMergeCount; j++) {
					temps.add(sheet.getMergedRegion(j));
				}
				this.rowReader.setCellRangeAddress(sheetIndex, temps);
				
				curRow=0;
				if (orderedBSRs == null) {
					orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
				}
				sheetName = orderedBSRs[sheetIndex].getSheetname();
			}
			break;

		case SSTRecord.sid:
			sstRecord = (SSTRecord) record;
			break;

		case BlankRecord.sid:
			BlankRecord brec = (BlankRecord) record;
			thisRow = brec.getRow();
			thisColumn = brec.getColumn();
			thisStr = "";
			// rowlist.add(thisColumn, thisStr);
			rowlist.put(thisColumn, thisStr);
			break;
		case BoolErrRecord.sid: // 单元格为布尔类型
			BoolErrRecord berec = (BoolErrRecord) record;
			thisRow = berec.getRow();
			thisColumn = berec.getColumn();
			thisStr = berec.getBooleanValue() + "";
			// rowlist.add(thisColumn, thisStr);
			rowlist.put(thisColumn, thisStr);
			break;

		case FormulaRecord.sid: // 单元格为公式类型
			FormulaRecord frec = (FormulaRecord) record;
			thisRow = frec.getRow();
			thisColumn = frec.getColumn();

			if (outputFormulaValues) {
				if (Double.isNaN(frec.getValue())) {
					// Formula result is a string
					// This is stored in the next record
					outputNextStringRecord = true;
					nextRow = frec.getRow();
					nextColumn = frec.getColumn();
				} else {
					thisStr = ((Double) frec.getValue()).toString();
					// thisStr = formatListener.formatNumberDateCell(frec);
				}
			} else {
				thisStr = '"' + HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression()) + '"';
			}
			// rowlist.add(thisColumn, thisStr);
			rowlist.put(thisColumn, thisStr);
			break;
		case StringRecord.sid:// 单元格中公式的字符串
			if (outputNextStringRecord) {
				// String for formula
				StringRecord srec = (StringRecord) record;
				thisStr = srec.getString();
				thisRow = nextRow;
				thisColumn = nextColumn;
				outputNextStringRecord = false;
			}
			break;
		case LabelRecord.sid:
			LabelRecord lrec = (LabelRecord) record;
			thisColumn = lrec.getColumn();
			value = lrec.getValue().trim();
			// this.rowlist.add(thisColumn, value);
			rowlist.put(thisColumn, value);
			break;
		case LabelSSTRecord.sid: // 单元格为字符串类型
			LabelSSTRecord lsrec = (LabelSSTRecord) record;
			thisColumn = lsrec.getColumn();
			if (sstRecord == null) {
				// rowlist.add(thisColumn, " ");
			} else {
				value = sstRecord.getString(lsrec.getSSTIndex()).toString().trim();
				// rowlist.add(thisColumn, value);
				rowlist.put(thisColumn, value);
			}
			break;
		case NumberRecord.sid: // 单元格为数字类型
			NumberRecord numrec = (NumberRecord) record;
			thisColumn = numrec.getColumn();
			value = ((Double) numrec.getValue()).toString();
			// value = formatListener.formatNumberDateCell(numrec).trim();
			// value = value.equals("") ? " " : value;
			// 向容器加入列值
			// rowlist.add(thisColumn, value);
			rowlist.put(thisColumn, value);
			break;
		default:
			break;
		}

		// 空值的操作
//		if (record instanceof MissingCellDummyRecord) {
//			MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
//			thisColumn = mc.getColumn();
//			rowlist.remove(thisColumn);
//		}
		// 行结束时的操作
		if (record instanceof LastCellOfRowDummyRecord) {
			if (rowlist.size() > 0) {
				// 每行结束时， 调用getRows() 方法
				rowReader.getRows(sheetIndex, curRow, rowlist);
				// 清空容器
				rowlist = new HashMap<Integer, String>();
			}
			curRow++;
		}
	}

	@Override
	public void close() throws WebServiceException {
		try {
			this.workbook.close();
		} catch (IOException e) {
		}
	}

}