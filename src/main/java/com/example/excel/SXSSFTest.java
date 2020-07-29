package com.example.excel;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SXSSFTest {
	public static void main(String[] args) {
		SXSSFTest t = new SXSSFTest();
		t.readTest("", "");

	}

	public void readTest(String filePath, String sheetName) {
		SXSSFWorkbook wb = createWb(filePath);
		Sheet sheet = wb.getSheet(sheetName);

		Row row = sheet.getRow(0);
		System.out.println("row=" + row);
		Cell cell1 = row.getCell(0);
		System.out.println(getStringFormulaValue(cell1));
		Cell cell2 = row.getCell(1);
		System.out.println(getStringFormulaValue(cell2));
		Cell cell3 = row.getCell(2);
		System.out.println(getStringFormulaValue(cell3));
		Cell cell4 = row.getCell(3);
		System.out.println(getStringFormulaValue(cell4));
		Cell cell5 = row.getCell(4);
		System.out.println(getStringFormulaValue(cell5));

	}

	public void writeTest(String filePath, String sheetName) {
		SXSSFWorkbook wb = createWb(filePath);
		Sheet sheet = wb.getSheet(sheetName);
	}

	private SXSSFWorkbook createWb(String filePath) {
		SXSSFWorkbook result = null;
		try {
			File file = new File(filePath);
			XSSFWorkbook xWb = (XSSFWorkbook) WorkbookFactory.create(new FileInputStream(file));
			result = new SXSSFWorkbook(xWb);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}

	private String getCellValue(Cell cell) {

		String result = "";
		if (cell != null) {
			switch (cell.getCellType()) {
			case STRING:
				result = cell.getStringCellValue();
				break;
			case NUMERIC:
				result = Double.toString(cell.getNumericCellValue());
				break;
			case BOOLEAN:
				result = Boolean.toString(cell.getBooleanCellValue());
				break;
			case FORMULA:
				result = getStringFormulaValue(cell);
				break;
//			case BLANK:
//				return getStringRangeValue(cell);
			default:
				System.out.println(cell.getCellType());
			}
		}
		return result;
	}

	/**
	 * 数式の計算結果を取得
	 * 
	 * @param cell セル
	 * @return 数式の計算結果
	 */
	private String getStringFormulaValue(Cell cell) {
		Workbook book = cell.getSheet().getWorkbook();
		CreationHelper helper = book.getCreationHelper();
		FormulaEvaluator evaluator = helper.createFormulaEvaluator();
		CellValue value = evaluator.evaluate(cell);
		switch (value.getCellType()) {
		case STRING:
			return value.getStringValue();
		case NUMERIC:
			return Double.toString(value.getNumberValue());
		case BOOLEAN:
			return Boolean.toString(value.getBooleanValue());
		default:
			System.out.println(value.getCellType());
			return null;
		}
	}

	/**
	 * キャッシュ値
	 * 
	 * @param cell
	 * @return
	 */
	public String getStringCachedFormulaValue(Cell cell) {
		switch (cell.getCachedFormulaResultType()) {
		case STRING:
			return cell.getStringCellValue();
		case NUMERIC:
			return Double.toString(cell.getNumericCellValue());
		case BOOLEAN:
			return Boolean.toString(cell.getBooleanCellValue());
		default:
			System.out.println(cell.getCachedFormulaResultType());
			return null;
		}
	}

//	public static String getStringRangeValue(Cell cell) {
//		int rowIndex = cell.getRowIndex();
//		int columnIndex = cell.getColumnIndex();
//
//		Sheet sheet = cell.getSheet();
//		int size = sheet.getNumMergedRegions();
//		for (int i = 0; i < size; i++) {
//			CellRangeAddress range = sheet.getMergedRegion(i);
//			if (range.isInRange(rowIndex, columnIndex)) {
//				Cell firstCell = getCell(sheet, range.getFirstRow(), range.getFirstColumn()); // 左上のセルを取得
//				return getStringValue(firstCell);
//			}
//		}
//		return null;
//	}
}
