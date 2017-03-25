package com.excel.util.tool;

import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import com.excel.util.intefaces.IRowReader;

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
public class Excel2007Reader2 extends DefaultHandler implements Closeable {
	// 表格数据类型
	private enum xssfDataType {
		BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER,
	}

	/**
	 * 表样式
	 */
	private StylesTable stylesTable;
	private xssfDataType nextDataType;
	// 用于格式化数值单元格的值。
	private short formatIndex;
	private String formatString;
	private final DataFormatter formatter = new DataFormatter();

	// 共享字符串表
	private SharedStringsTable sst;

	private ReadOnlySharedStringsTable sharedStringsTable;

	private StringBuffer value = new StringBuffer();

	// 当前所处的列
	private int thisColumn = -1;

	private int sheetIndex = -1;
	private Map<Integer, String> rowlist = new HashMap<Integer, String>();
	// 当前行
	private int curRow = 0;

	private boolean vIsOpen;

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
		if (this.pkg != null) {
			this.pkg.close();
		}
		this.pkg = OPCPackage.open(stream);
		XSSFReader r = new XSSFReader(pkg);
		this.stylesTable = r.getStylesTable();
		this.sharedStringsTable = new ReadOnlySharedStringsTable(this.pkg);
		SharedStringsTable sstTemp = r.getSharedStringsTable();
		XMLReader parser = fetchSheetParser(sstTemp);
		Iterator<InputStream> sheets = r.getSheetsData();
		while (sheets.hasNext()) {
			curRow = 0;
			sheetIndex++;
			InputStream sheet = sheets.next();
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
		}
	}

	public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {

		if ("inlineStr".equals(name) || "v".equals(name)) {
			vIsOpen = true;
			value.setLength(0);
		}

		// c => 单元格
		if ("c".equals(name)) {
			String r = attributes.getValue("r");
			int firstDigit = -1;
			for (int c = 0; c < r.length(); ++c) {
				if (Character.isDigit(r.charAt(c))) {
					firstDigit = c;
					break;
				}
			}
			// 当前列
			this.thisColumn = nameToColumn(r.substring(0, firstDigit));

			this.nextDataType = xssfDataType.NUMBER;
			this.formatIndex = -1;
			this.formatString = null;

			String cellType = attributes.getValue("t");
			String cellStyleStr = attributes.getValue("s");
			if ("b".equals(cellType))
				nextDataType = xssfDataType.BOOL;
			else if ("e".equals(cellType))
				nextDataType = xssfDataType.ERROR;
			else if ("inlineStr".equals(cellType))
				nextDataType = xssfDataType.INLINESTR;
			else if ("s".equals(cellType))
				nextDataType = xssfDataType.SSTINDEX;
			else if ("str".equals(cellType))
				nextDataType = xssfDataType.FORMULA;
			else if (cellStyleStr != null) {
				int styleIndex = Integer.parseInt(cellStyleStr);
				XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
				this.formatIndex = style.getDataFormat();
				this.formatString = style.getDataFormatString();
				if (this.formatString == null)
					this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
			}
		}
	}

	public void endElement(String uri, String localName, String name) throws SAXException {
		String thisStr = null;
		System.out.println("name  " +name + "  value =" + value);
		if ("v".equals(name)) {
			switch (nextDataType) {

			case BOOL:
				char first = value.charAt(0);
				thisStr = first == '0' ? "FALSE" : "TRUE";
				break;

			case ERROR:
				thisStr = "\"ERROR:" + value.toString() + '"';
				break;

			case FORMULA:
				thisStr = '"' + value.toString() + '"';
				break;

			case INLINESTR:
				XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
				thisStr = '"' + rtsi.toString() + '"';
				break;

			case SSTINDEX:
				String sstIndex = value.toString();
				try {
					int idx = Integer.parseInt(sstIndex);
					XSSFRichTextString rtss = new XSSFRichTextString(sharedStringsTable.getEntryAt(idx));
					thisStr = '"' + rtss.toString() + '"';
				} catch (NumberFormatException ex) {
					System.out.println("Failed to parse SST index '" + sstIndex + "': " + ex.toString());
				}
				break;

			case NUMBER:
				String n = value.toString();
				thisStr = n;
				// if (HSSFDateUtil.isADateFormat(this.formatIndex, n)) {
				// Double d = Double.parseDouble(n);
				// Date date = HSSFDateUtil.getJavaDate(d);
				// thisStr = formateDateToString(date);
				// } else if (this.formatString != null) {
				// thisStr =
				// formatter.formatRawCellContents(Double.parseDouble(n),
				// this.formatIndex,
				// this.formatString);
				// } else {
				// thisStr = n;
				// }
				break;

			default:
				thisStr = "(TODO: Unexpected type: " + nextDataType + ")";
				break;
			}
			rowlist.put(thisColumn, thisStr);
		} else if ("row".equals(name)) {
			if (rowlist.size() > 0) {
				rowReader.getRows(sheetIndex, curRow, rowlist);
				rowlist = new HashMap<Integer, String>();
			}
			curRow++;
		}
		
	}

	// 创建XML访问对象
	private XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException, ParserConfigurationException {
		SAXParserFactory saxFactory = SAXParserFactory.newInstance();
		SAXParser saxParser = saxFactory.newSAXParser();
		XMLReader parser = saxParser.getXMLReader();
		this.sst = sst;
		parser.setContentHandler(this);
		return parser;
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

	// 时间类型转换
	private String formateDateToString(Date date) {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");// 格式化日期
		return sdf.format(date);

	}

	public void characters(char[] ch, int start, int length) throws SAXException {
		if (vIsOpen)
			value.append(ch, start, length);
	}

	@Override
	public void close() throws IOException {
		this.pkg.close();
	}
}
