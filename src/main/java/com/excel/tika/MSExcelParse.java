package com.excel.tika;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.tika.exception.TikaException;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.microsoft.ooxml.OOXMLParser;
import org.apache.tika.sax.BodyContentHandler;
import org.xml.sax.SAXException;

public class MSExcelParse {
	public static void main(final String[] args) {

		try {
			MSExcelParse msp = new MSExcelParse();
			msp.parse("test.xlsx");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void parse(String fileName) throws IOException, TikaException, SAXException {
		BodyContentHandler handler = new BodyContentHandler();
		Metadata metadata = new Metadata();
		FileInputStream inputstream = new FileInputStream(new File(fileName));
		ParseContext pcontext = new ParseContext();

		OOXMLParser msofficeparser = new OOXMLParser();
		msofficeparser.parse(inputstream, handler, metadata, pcontext);
		System.out.println("Contents of the document:" + handler.toString());
		System.out.println("Metadata of the document:");
		String[] metadataNames = metadata.names();

		for (String name : metadataNames) {
			System.out.println(name + ": " + metadata.get(name));
		}
	}
}
