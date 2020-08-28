/**
 * 
 */
package edu.abhi.poi.excel;

import java.util.concurrent.Callable;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * @author abhishek sarkar
 *
 */
public class SheetProcessorTask implements Callable<CallableValue> {

	private XSSFSheet sheet1, sheet2;
	private boolean remarksOnly;
	private CallableValue crt;

	public SheetProcessorTask(XSSFSheet sheet1, XSSFSheet sheet2, boolean commentFlag) {
		this.sheet1 = sheet1;
		this.sheet2 = sheet2;
		this.remarksOnly = commentFlag;
	}

	@Override
	public CallableValue call() {
		crt = new CallableValue();
		try {
			processAllRows();
		} catch (Exception e) {
			crt.setFailed(true);
			crt.setException(e);
		}
		return crt;
	}

	private void processAllRows() throws Exception {
		for (int rowIndex = 0; rowIndex <= sheet1.getLastRowNum(); rowIndex++) {
			XSSFRow row1 = (XSSFRow) sheet1.getRow(rowIndex);
			XSSFRow row2 = (XSSFRow) sheet2.getRow(rowIndex);

			if (row1 == null || row2 == null) {
				if (!(row1 == null && row2 == null)) {
					crt.setDiffFlag(true);
					processNullRow(sheet1, rowIndex, row2);
				}
				continue;
			}

			processAllColumns(row1, row2);
		}
	}

	private void processAllColumns(XSSFRow row1, XSSFRow row2) throws Exception {
		StringBuilder rowRemarks = new StringBuilder();
		boolean isRow1Blank = true, isRow2Blank = true;
		
		for (int columnIndex = 0; columnIndex <= row1.getLastCellNum(); columnIndex++) {
			XSSFCell cell1 = (XSSFCell) row1.getCell(columnIndex);
			XSSFCell cell2 = (XSSFCell) row2.getCell(columnIndex);
			
			if (Utility.hasNoContent(cell1)) {
				if (Utility.hasContent(cell2)) {
					isRow2Blank = false;
					crt.setDiffFlag(true);
					Utility.processDiffForColumn(cell1 == null ? row1.createCell(columnIndex) : cell1, remarksOnly,
							Utility.getCellValue(cell2), rowRemarks);
				}
			} else if (Utility.hasNoContent(cell2)) {
				if (Utility.hasContent(cell1)) {
					isRow1Blank = false;
					crt.setDiffFlag(true);
					Utility.processDiffForColumn(cell1, remarksOnly, cell2 == null? null : Utility.getCellValue(cell2), rowRemarks);
				}
			} else {
				isRow1Blank = isRow2Blank = false;
				
				if (!Utility.getCellValue(cell1).equals(Utility.getCellValue(cell2))) {
					crt.setDiffFlag(true);
					Utility.processDiffForColumn(cell1, remarksOnly, Utility.getCellValue(cell2), rowRemarks);
				}
			}
		}
		
		if(!isRow1Blank && isRow2Blank)
			crt.getDiffContainer().append(String.format("\nRemoved Row[%s] of Sheet[%s]",
					(row1.getRowNum() + 1), sheet1.getSheetName()));
		else if(isRow1Blank && !isRow2Blank)
			crt.getDiffContainer().append(String.format("\nAdded Row[%s] in Sheet[%s]",
					(row1.getRowNum() + 1), sheet1.getSheetName()));
		else
			crt.getDiffContainer().append(rowRemarks);
	}

	public void processNullRow(XSSFSheet sheet1, int rowIndex, XSSFRow row2) throws Exception {
		XSSFRow row1 = sheet1.getRow(rowIndex);
		StringBuilder rowRemarks = new StringBuilder();
		
		if (row1 == null) {
			if (row2.getPhysicalNumberOfCells() != 0) {
				row1 = sheet1.createRow(rowIndex);
				crt.setDiffFlag(true);
				for (int columnIndex = 0; columnIndex <= row2.getLastCellNum(); columnIndex++) {
					Utility.processDiffForColumn(row1.createCell(0), remarksOnly,
							Utility.getCellValue(row2.getCell(columnIndex)), rowRemarks);
				}
				crt.getDiffContainer().append(String.format("\nAdded Row[%s] in Sheet[%s]",
						(row1.getRowNum() + 1), sheet1.getSheetName()));
			}
		} else {
			if (row1.getPhysicalNumberOfCells() != 0) {
				crt.setDiffFlag(true);
				XSSFCell cell1 = row1.getCell(0);
				Utility.processDiffForColumn(cell1 == null ? row1.createCell(0) : cell1, remarksOnly, "Null row", rowRemarks);
				crt.getDiffContainer().append(String.format("\nRemoved Row[%s] of Sheet[%s]",
						(row1.getRowNum() + 1), sheet1.getSheetName()));
			}
		}
	}

}
