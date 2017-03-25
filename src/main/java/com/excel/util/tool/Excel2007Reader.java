package com.excel.util.tool;

import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.util.CellRangeAddressBase;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import com.excel.util.intefaces.IRowReader;
import com.excel.util.model.ExcelCellRangeAddress;

/**
 * 抽象Excel2007读取器，excel2007的底层数据结构是xml文件，采用SAX的事件驱动的方法解析
 * xml，需要继承DefaultHandler，在遇到文件内容时，事件会触发，这种做法可以大大降低 内存的耗费，特别使用于大数据量的文件。
 *
 */
/**
 * 
 * @author Bless
 * @version 1.0
 */
public class Excel2007Reader extends DefaultHandler implements Closeable {
	
	private Map<Integer, List<CellRangeAddressBase>> mergeCells; 
	
	private Character[] numChar = new Character[]{'0','1','2','3','4','5','6','7','8','9'};
	
	private boolean isTElement;
	private String lastContents;
	// 当前所处的列
	private int thisColumn = -1;

	private int sheetIndex = -1;
	private Map<Integer, String> rowlist = new HashMap<Integer, String>();
	// 当前行
	private int curRow = -1;
	private SharedStringsTable sst;
	private IRowReader rowReader;

	private OPCPackage pkg;

	public void setRowReader(IRowReader rowReader) {
		this.rowReader = rowReader;
	}

	/**
	 * 遍历工作簿中所有的电子表格
	 * 
	 * @param filename
	 * @throws Exception
	 */
	public void process(InputStream stream) throws Exception {
		mergeCells = new HashMap<>();
		if (this.pkg != null) {
			this.pkg.close();
		}
		this.pkg = OPCPackage.open(stream);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sstTemp = r.getSharedStringsTable();
		XMLReader parser = fetchSheetParser(sstTemp);
		Iterator<InputStream> sheets = r.getSheetsData();
		while (sheets.hasNext()) {
			sheetIndex++;
			if(!mergeCells.containsKey(sheetIndex)){
				mergeCells.put(sheetIndex, new ArrayList<>());
			}
			InputStream sheet = sheets.next();
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
		}
		if(this.rowReader != null && mergeCells!= null){
			Set<Integer> keys = this.mergeCells.keySet();
			if(keys != null && keys.size() > 0){
				Integer[] sheetIndex = keys.toArray(new Integer[keys.size()]);
				Arrays.sort(sheetIndex);
				List<List<CellRangeAddressBase>> cellRangeAddressBases = new ArrayList<>();
				for (Integer item : sheetIndex) {
					cellRangeAddressBases.add(mergeCells.get(item));
					this.rowReader.setCellRangeAddress(item, mergeCells.get(item));
				}
			}
		}
	}
	public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException { 
		if(name.equals("mergeCells")){
			String count = attributes.getValue("count");
			//合并总数
		} else if(name.equals("mergeCell")){
			String ref = attributes.getValue("ref");
			//分析合并
			List<CellRangeAddressBase> mergeCellList = mergeCells.get(sheetIndex);
			String[] strArr = ref.split(":");
			int firstCol = 0;
			int firstRow = 0;
			int lastCol = 0;
			int lastRow = 0;
			for (int i = 0; i < strArr.length; i++) {
				String str = strArr[i];
				String colName = "";
				String rowIndex = "";
				for (int j = 0; j < str.length(); j++) {
					Character c = str.charAt(j);
					Boolean state = false;
					for (Character numC : numChar) {
						if(numC.equals(c)) state = true;
					}
					if(state){
						rowIndex = rowIndex + c;
					}else {
						colName = colName + c;
					}
				}
				if(i == 0){
					firstCol = nameToColumn(colName);
					firstRow = Integer.valueOf(rowIndex) - 1;
				} else{
					lastCol = nameToColumn(colName);
					lastRow = Integer.valueOf(rowIndex) - 1;
				}
			}
			ExcelCellRangeAddress cellRangeAddress = new ExcelCellRangeAddress(firstRow,lastRow,firstCol,lastCol);
			mergeCellList.add(cellRangeAddress);
		}
		if ("c".equals(name)) {
			String r = attributes.getValue("r");
			int firstDigit = -1;
			for (int c = 0; c < r.length(); ++c) {
				if (Character.isDigit(r.charAt(c))) {
					firstDigit = c;
					break;
				}
			}
			this.thisColumn = nameToColumn(r.substring(0, firstDigit));
			String colIndex = r.substring(firstDigit);
			if(colIndex!=null && !colIndex.equals("")){
				this.curRow = Integer.valueOf(colIndex) - 1;
			}
			String cellType = attributes.getValue("t");
			if (cellType != null && cellType.equals("s")) {
				isTElement = true;
			} else {
				isTElement = false;
			}
		}
        lastContents = ""; 
	}

	public void endElement(String uri, String localName, String name) throws SAXException {
		
		
		if(isTElement){
			try {
				int idx = Integer.parseInt(lastContents);
				lastContents = new XSSFRichTextString(sst.getEntryAt(idx))
						.toString();
			} catch (Exception e) {

			}
        }
		if ("v".equals(name) || "t".equals(name) || "s".equals(name)) {
			String value = lastContents.trim();  
            value = value.equals("")?" ":value;  
			rowlist.put(thisColumn, value);
		} else if ("row".equals(name)) {
			if (rowlist.size() > 0) {
				rowReader.getRows(sheetIndex, curRow, rowlist);
				rowlist = new HashMap<Integer, String>();
			}
		}
		
	}
	
	// 获取列的下标
	private int nameToColumn(String name) {
		int column = -1;
		for (int i = 0; i < name.length(); ++i) {
			int c = name.charAt(i);
			column = (column + 1) * 26 + c - 'A';
		}
		return column;
	}

	// 创建XML访问对象
	private XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException, ParserConfigurationException {
		SAXParserFactory saxFactory = SAXParserFactory.newInstance();
		SAXParser saxParser = saxFactory.newSAXParser();
		XMLReader parser = saxParser.getXMLReader();
		parser.setContentHandler(this);
		this.sst = sst;
		return parser;
	}

	public void characters(char[] ch, int start, int length) throws SAXException {
		//得到单元格内容的值  
        lastContents += new String(ch, start, length);  
	}

	@Override
	public void close() throws IOException {
		this.pkg.close();
	}
}
