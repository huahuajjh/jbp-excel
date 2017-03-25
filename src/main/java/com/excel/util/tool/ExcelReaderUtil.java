package com.excel.util.tool;

import java.io.InputStream;
import java.io.PushbackInputStream;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import com.excel.util.intefaces.IRowReader;

/**
 * 
 * @author Bless
 * @version 1.0
 */
public class ExcelReaderUtil {
	/**
	 * 读取Excel文件，可能是03也可能是07版本
	 * 
	 * @param excel03
	 * @param excel07
	 * @param fileName
	 * @throws Exception
	 */
	public static void readExcel(IRowReader reader, InputStream stream) throws Exception {
		if (!stream.markSupported()) {
			PushbackInputStream inputStream = new PushbackInputStream(stream, 8);
			if (POIFSFileSystem.hasPOIFSHeader(inputStream)) {// 处理excel2003文件
				Excel2003Reader excel03 = new Excel2003Reader();
				excel03.setRowReader(reader);
				excel03.process(inputStream);
				excel03.close();
				return;
			} else if (POIXMLDocument.hasOOXMLHeader(inputStream)) {// 处理excel2007文件
				Excel2007Reader excel07 = new Excel2007Reader();
				excel07.setRowReader(reader);
				excel07.process(inputStream);
				excel07.close();
				return;
			}
		}
		throw new Exception("文件格式错误，fileName的扩展名只能是xls或xlsx。");

	}
}