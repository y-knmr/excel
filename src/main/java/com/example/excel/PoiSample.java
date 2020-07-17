package com.example.excel;

import java.io.File;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class PoiSample {
	public static void main(String[] args) {

		String fileName = "test.xlsx";
		try (Workbook wb = WorkbookFactory.create(new File(fileName))) {

			System.out.println(wb.getNumberOfSheets());

			wb.sheetIterator().forEachRemaining(s -> {
				System.out.print(String.format("%s : %s : ", s.getSheetName(), s.getClass().getName()));

//				if (s instanceof HSSFSheet) {
//					HSSFSheet sheet = (HSSFSheet) s;
//					try {
//						System.out.println(sheet.getDialog());
//					} catch (NullPointerException e) {
//						e.printStackTrace();
//					}
//				} else {
//					System.out.println();
//				}
			});
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
