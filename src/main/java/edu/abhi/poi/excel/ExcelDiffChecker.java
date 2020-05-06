package edu.abhi.poi.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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

	public static void main(String[] args) {

		String FILE_NAME1 = args[0];
		String FILE_NAME2 = args[1];

		String RESULT_FILE = FILE_NAME1.substring(0, FILE_NAME1.lastIndexOf(".")) + " vs " +
				FILE_NAME2;

		boolean success = true;

		File resultFile = new File(RESULT_FILE);
		if(resultFile.exists())
			resultFile.delete();

		try(XSSFWorkbook resultWorkbook = new XSSFWorkbook(new FileInputStream(new File(FILE_NAME1)));
				XSSFWorkbook workbook2 = new XSSFWorkbook(new FileInputStream(new File(FILE_NAME2)))) {

			if(resultWorkbook.getNumberOfSheets() != workbook2.getNumberOfSheets()) {
				System.out.println("Unequal number of sheets present!");
				success = false;
			}

			breaker_on_failure:
				for(int sheetIndex = 0; (sheetIndex < resultWorkbook.getNumberOfSheets()) && success; sheetIndex++) {
					XSSFSheet sheet1 = (XSSFSheet) resultWorkbook.getSheetAt(sheetIndex);
					XSSFSheet sheet2 = (XSSFSheet) workbook2.getSheetAt(sheetIndex);

					if(!sheet1.getSheetName().equals(sheet2.getSheetName())) {
						System.out.println(String.format("Expected sheet name[%s] of workbook[%s] but found [%s]!",
								sheet1.getSheetName(), FILE_NAME2, sheet2.getSheetName()));
						success = false;
						break breaker_on_failure;
					} else if(sheet1.getLastRowNum() != sheet2.getLastRowNum()) {
						System.out.println(String.format("Expected row count[%s] of sheet[%s] of workbook[%s] but found [%s]!",
								sheet1.getLastRowNum(), sheet1.getSheetName(), FILE_NAME2, sheet2.getLastRowNum()));
						success = false;
						break breaker_on_failure;
					}

					for(int rowIndex = 0; rowIndex < sheet1.getLastRowNum(); rowIndex++) {
						XSSFRow row1 = (XSSFRow) sheet1.getRow(rowIndex);
						XSSFRow row2 = (XSSFRow) sheet2.getRow(rowIndex);
						
						if(row1 == null || row2 == null) {
							if(row1 != row2) 
								System.out.println("Both rows are not null at rowIndex = " + rowIndex);
							continue;
						} else if(row1.getLastCellNum() != row2.getLastCellNum()) {
							System.out.println(String.format("Expected column count[%s] of rowIndex[%s] of sheet[%s] of workbook[%s] but found [%s]!",
									row1.getLastCellNum(), rowIndex, sheet1.getSheetName(), FILE_NAME2, row2.getLastCellNum()));
							success = false;
							break breaker_on_failure;
						}

						for(int columnIndex = 0; columnIndex < row1.getLastCellNum(); columnIndex++) {
							XSSFCell cell1 = (XSSFCell) row1.getCell(columnIndex);
							XSSFCell cell2 = (XSSFCell) row2.getCell(columnIndex);

							if(hasNoContent(cell1)) {
								if(hasContent(cell2)) {
									if(cell1 == null)
										cell1 = row1.createCell(columnIndex);
									System.out.println(String.format("Diff at cell[%s] of sheet[%s]", 
											cell1.getReference(), sheet1.getSheetName()));
									
									addComment(resultWorkbook, sheet1, rowIndex, cell1, "SYSTEM", cell2);
								}
							} else if(hasNoContent(cell2)) {
								if(hasContent(cell1)) {
									System.out.println(String.format("Diff at cell[%s] of sheet[%s]", 
											cell1.getReference(), sheet1.getSheetName()));

									addComment(resultWorkbook, sheet1, rowIndex, cell1, "SYSTEM", null);
								}
							} else if(!cell1.getRawValue().equals(cell2.getRawValue())) {
								System.out.println(String.format("Diff at cell[%s] of sheet[%s]", 
										cell1.getReference(), sheet1.getSheetName()));
								
								addComment(resultWorkbook, sheet1, rowIndex, cell1, "SYSTEM", cell2);
							}
						}
					}
				}
	        
	        try (FileOutputStream outputStream = new FileOutputStream(RESULT_FILE)) {
	        	resultWorkbook.write(outputStream);
	        }
		} catch (Exception e) {
			e.printStackTrace(System.out);
			success = false;
		} finally {
			if(success)
				System.out.println("Diff excel has been generated!");
			else
				resultFile.delete();
		}

	}
	
	private static boolean hasNoContent(XSSFCell cell) {
		return cell == null || cell.getRawValue() == null || cell.getRawValue().equals("");
	}
	
	private static boolean hasContent(XSSFCell cell) {
		return cell != null && cell.getRawValue() != null && !cell.getRawValue().equals("");
	}
	
	@SuppressWarnings("rawtypes")
	
	private static void addComment(Workbook workbook, Sheet sheet, int rowIndex, Cell cell1, String author, XSSFCell cell2) throws Exception {
		CreationHelper factory = workbook.getCreationHelper();
        //get an existing cell or create it otherwise:

        ClientAnchor anchor = factory.createClientAnchor();
        //i found it useful to show the comment box at the bottom right corner
        anchor.setCol1(cell1.getColumnIndex() + 1); //the box of the comment starts at this given column...
        anchor.setCol2(cell1.getColumnIndex() + 3); //...and ends at that given column
        anchor.setRow1(rowIndex + 1); //one row below the cell...
        anchor.setRow2(rowIndex + 5); //...and 4 rows high

        Drawing drawing = sheet.createDrawingPatriarch();
        Comment comment = drawing.createCellComment(anchor);
        
        //set the comment text and author
        comment.setString(factory.createRichTextString("Found " + getCellValue(cell2)));
        comment.setAuthor(author);

        cell1.setCellComment(comment);
    }
	
	private static String getCellValue(XSSFCell cell2) throws Exception {
		String content = "";
		
		CellType cellType = cell2.getCellType();
		
		if(cellType == CellType.FORMULA)
			cellType = cell2.getCachedFormulaResultType();
		
		switch(cellType) {
		case BLANK:	content += null;
			break;
		case BOOLEAN: content += cell2.getBooleanCellValue();
			break;
		case ERROR: content += cell2.getErrorCellString();
			break;
		case STRING: content += cell2.getRichStringCellValue();
			break;
		case NUMERIC: content += cell2.getNumericCellValue();
			break;
		case _NONE:	content += null;
			break;
		default: throw new Exception("Unexpected Cell Type = " + cell2.getCellType());
        }
		return content;
	}
	
}