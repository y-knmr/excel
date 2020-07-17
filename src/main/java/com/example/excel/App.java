package com.example.excel;

import java.util.List;
import java.util.Map;

/**
 * Hello world!
 *
 */
public class App {
	public static void main(String[] args) {
		//String filename = "sample.xlsm";
		String filename = "test.xlsx";
		XssfSax xs = new XssfSax();

		try {
			List<Map<Integer, String>> ret = xs.read(filename, "ふつうのシート");
			for (Map<Integer, String> map : ret) {
				System.out.println(map);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
