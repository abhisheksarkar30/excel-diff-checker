/**
 * 
 */
package edu.abhi.poi.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;

/**
 * @author abhis
 *
 */
public class Utility {

	public static boolean hasNoContent(XSSFCell cell) {
		return cell == null || cell.getRawValue() == null || cell.getRawValue().equals("");
	}

	public static boolean hasContent(XSSFCell cell) {
		return cell != null && cell.getRawValue() != null && !cell.getRawValue().equals("");
	}

	@SuppressWarnings("rawtypes")
	public static void addComment(Workbook workbook, Sheet sheet, int rowIndex, XSSFCell cell1, String author, XSSFCell cell2) throws Exception {
		
		System.out.println(String.format("Diff at cell[%s] of sheet[%s]", cell1.getReference(), sheet.getSheetName()));
		
		ExcelDiffChecker.diffFound = true;
		
		if(!ExcelDiffChecker.commentFlag) {
			System.out.println(String.format("Expected: [%s], Found: [%s]", getCellValue(cell1), getCellValue(cell2)));
			return;
		}
		
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
	    comment.setString(factory.createRichTextString("Found " + Utility.getCellValue(cell2)));
	    comment.setAuthor(author);
	
	    cell1.setCellComment(comment);
	}

	public static String getCellValue(XSSFCell cell) throws Exception {
		String content = "";
		
		CellType cellType = cell.getCellType();
		
		if(cellType == CellType.FORMULA)
			cellType = cell.getCachedFormulaResultType();
		
		switch(cellType) {
		case BLANK:	content += null;
			break;
		case BOOLEAN: content += cell.getBooleanCellValue();
			break;
		case ERROR: content += cell.getErrorCellString();
			break;
		case STRING: content += cell.getRichStringCellValue();
			break;
		case NUMERIC: content += DateUtil.isCellDateFormatted(cell) ? cell.getDateCellValue() : 
			cell.getNumericCellValue();
			break;
		case _NONE:	content += null;
			break;
		default: throw new Exception(String.format("Unexpected Cell[%s] Type[%s] of sheet[%s]", cell.getReference(), 
				cell.getCellType(), cell.getSheet().getSheetName()));
	    }
		return content;
	}

}
