package com.example.excel;

import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

public class XssfSax {

	private HashMap<String, String> workMap = null;
	private String workKey = "";

	public List<Map<Integer, String>> read(String fileName, String sheetName) throws Exception {
		workMap = new HashMap<>();
		List<Map<Integer, String>> ret = new ArrayList<>();

		OPCPackage pkg = OPCPackage.open(fileName, PackageAccess.READ);

		try {
			XSSFReader r = new XSSFReader(pkg);
			SharedStringsTable sst = r.getSharedStringsTable();

			XMLReader parser = fetchSheetParser(sst);

			Iterator<InputStream> sheets = r.getSheetsData();
			System.out.println("sheets=" + sheets);
			while (sheets.hasNext()) {
				InputStream sheet = sheets.next();
				if (sheetName.equals(((XSSFReader.SheetIterator) sheets).getSheetName())) {
					InputSource sheetSource = new InputSource(sheet);
					parser.parse(sheetSource);
					sheet.close();
				}
			}
		} finally {
			pkg.close();
		}

		for (String colRow : workMap.keySet()) {

			int sepIdx = 0;
			int colRowLen = colRow.length();

			for (int i = 0; i < colRowLen; i++) {
				if (colRow.charAt(i) < 'A') {
					sepIdx = i;
					break;
				}
			}

			String colStr = colRow.substring(0, sepIdx);
			String rowStr = colRow.substring(sepIdx);
			int x = 0;
			int y = Integer.parseInt(rowStr);
			int colStrLen = colStr.length();
			for (int i = 0; i < colStrLen; i++) {
				int colIdx = colStr.charAt(i) - 'A' + 1;
				for (int j = 1; j < colStrLen - i; j++) {
					colIdx = colIdx * 26;
				}
				x += colIdx;
			}
			x--;
			y--;

			Map<Integer, String> map = null;
			int retSize = ret.size();
			if (y < retSize) {
				map = ret.get(y);
			} else {
				int addCnt = y - retSize + 1;
				for (int i = 0; i < addCnt; i++) {
					map = new HashMap<>();
					ret.add(map);
				}
			}
			map.put(x, workMap.get(colRow));
		}

		workMap.clear();
		workMap = null;
		return ret;
	}

	private XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
		XMLReader parser = XMLReaderFactory.createXMLReader();
		ContentHandler handler = new SheetHandler(sst);
		parser.setContentHandler(handler);
		return parser;
	}

	private class SheetHandler extends DefaultHandler {
		private SharedStringsTable sst;
		private String lastContents;
		private String cellType;
		private boolean nextIsString;
		private boolean inlineStr;

		private SheetHandler(SharedStringsTable sst) {
			this.sst = sst;
		}

		@Override
		public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
			// c => cell
			if (name.equals("c")) {
				workKey = attributes.getValue("r");
				// Print the cell reference
				// System.out.print(attributes.getValue("r") + " - ");
				// Figure out if the value is an index in the SST
				cellType = attributes.getValue("t");
				nextIsString = cellType != null && cellType.equals("s");
				inlineStr = cellType != null && cellType.equals("inlineStr");
			}
			// Clear contents cache
			lastContents = "";
		}

		@Override
		public void endElement(String uri, String localName, String name) throws SAXException {
			// Process the last contents as required.
			// Do now, as characters() may be called more than once
			if (nextIsString) {
				int idx = Integer.parseInt(lastContents);
				lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
				nextIsString = false;
			}

			// v => contents of a cell
			// Output after we've seen the string contents
			if (name.equals("v") || (inlineStr && name.equals("c"))) {

				// 0.14が0.14000000000000001になる場合を考慮
				if (lastContents.contains(".") && cellType == null) {
					try {
						double d = Double.parseDouble(lastContents);
						lastContents = BigDecimal.valueOf(d).toPlainString();
					} catch (Exception e) {
						e.printStackTrace();
					}
				}
				// System.out.println(lastContents);
				workMap.put(workKey, lastContents);
			}
		}

		@Override
		public void characters(char[] ch, int start, int length) throws SAXException { // NOSONAR
			lastContents += new String(ch, start, length);
		}
	}
}
