package edu.abhi.poi.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.function.Consumer;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 
 * @author abhishek sarkar
 *
 */
public class ExcelDiffChecker {

	private static boolean diffFound = false;
	private static boolean commentFlag = true;
	private static String FILE_NAME1, FILE_NAME2;

	public static void main(String[] args) {

		FILE_NAME1 = args[0];
		FILE_NAME2 = args[1];
		commentFlag = args.length == 2;

		String RESULT_FILE = FILE_NAME1.substring(0, FILE_NAME1.lastIndexOf(".")) + " vs " + FILE_NAME2;

		File resultFile = new File(RESULT_FILE);

		Utility.deleteIfExists(resultFile);

		try(XSSFWorkbook resultWorkbook = new XSSFWorkbook(new FileInputStream(new File(FILE_NAME1)));
				XSSFWorkbook workbook2 = new XSSFWorkbook(new FileInputStream(new File(FILE_NAME2)))) {

			processAllSheets(resultWorkbook, workbook2);

			if(diffFound) {
				if(commentFlag) {
					try (FileOutputStream outputStream = new FileOutputStream(RESULT_FILE)) {
						resultWorkbook.write(outputStream);
						System.out.println("Diff excel has been generated!");
					}
				}
			} else 
				System.out.println("No diff found!");
		} catch (Exception e) {
			e.printStackTrace(System.out);
		}
	}

	private static void processAllSheets(XSSFWorkbook resultWorkbook, XSSFWorkbook workbook2) throws Exception {

		Consumer<Sheet> consumer = sheet1 -> {
			XSSFSheet sheet2 = (XSSFSheet) workbook2.getSheet(sheet1.getSheetName());

			if(sheet2 == null) {
				System.out.println(String.format("Sheet[%s] doesn't exist in workbook[%s]", sheet1.getSheetName(), FILE_NAME2));
			} else
				try {
					processAllRows((XSSFSheet) sheet1, sheet2);
				} catch (Exception e) {
					e.printStackTrace(System.out);
				}
		};

		resultWorkbook.forEach(consumer);
	}

	private static void processAllRows(XSSFSheet sheet1, XSSFSheet sheet2) throws Exception {
		for(int rowIndex = 0; rowIndex <= sheet1.getLastRowNum(); rowIndex++) {
			XSSFRow row1 = (XSSFRow) sheet1.getRow(rowIndex);
			XSSFRow row2 = (XSSFRow) sheet2.getRow(rowIndex);

			if(row1 == null || row2 == null) {
				if(!(row1 == null && row2 ==null)) {
					diffFound = true;
					processNullRow(sheet1, rowIndex, row2);
				}
				continue;
			}

			processAllColumns(row1, row2);
		}
	}

	private static void processAllColumns(XSSFRow row1, XSSFRow row2) throws Exception {
		for(int columnIndex = 0; columnIndex <= row1.getLastCellNum(); columnIndex++) {
			XSSFCell cell1 = (XSSFCell) row1.getCell(columnIndex);
			XSSFCell cell2 = (XSSFCell) row2.getCell(columnIndex);

			if(Utility.hasNoContent(cell1)) {
				if(Utility.hasContent(cell2)) {
					diffFound = true;
					Utility.processDiffForColumn(cell1 == null? row1.createCell(columnIndex) : cell1, commentFlag, Utility.getCellValue(cell2));
				}
			} else if(Utility.hasNoContent(cell2)) {
				if(Utility.hasContent(cell1)) {
					diffFound = true;
					Utility.processDiffForColumn(cell1, commentFlag, Utility.getCellValue(cell2));
				}
			} else if(!cell1.getRawValue().equals(cell2.getRawValue())) {
				diffFound = true;
				Utility.processDiffForColumn(cell1, commentFlag, Utility.getCellValue(cell2));
			}
		}
	}

	public static void processNullRow(XSSFSheet sheet1, int rowIndex, XSSFRow row2) throws Exception {
		XSSFRow row1 = sheet1.getRow(rowIndex);

		if(row1 == null) {
			row1 = sheet1.createRow(rowIndex);

			for(int columnIndex = 0; columnIndex <= row2.getLastCellNum(); columnIndex++) {
				Utility.processDiffForColumn(row1.createCell(0), commentFlag, Utility.getCellValue(row2.getCell(columnIndex)));
			}
		} else {
			XSSFCell cell1 = row1.getCell(0);
			Utility.processDiffForColumn(cell1 == null? row1.createCell(0) : cell1, commentFlag, "Null row");
		}
	}

}