/**
 * 
 */
package edu.abhi.poi.excel;

import java.io.File;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;

/**
 * @author abhishek sarkar
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
	public static void processDiffForColumn(XSSFCell cell1, boolean remarksOnly, String note, StringBuilder sb) throws Exception {

		Sheet sheet = cell1.getSheet();

		if(remarksOnly) {
			sb.append(String.format("\nDiff at Cell[%s] of Sheet[%s]", cell1.getReference(), sheet.getSheetName()));
			sb.append(String.format("\nExpected: [%s], Found: [%s]", getCellValue(cell1), note));
			return;
		} else {
			sb.append(sb.length() > 0 ? " " : String.format("\nDiff at Row[%s] of Sheet[%s]: ", (cell1.getRowIndex() + 1), sheet.getSheetName()));
			sb.append(cell1.getReference());
		}

		synchronized(sheet.getWorkbook()) {
			CreationHelper factory = sheet.getWorkbook().getCreationHelper();
			//get an existing cell or create it otherwise:

			ClientAnchor anchor = factory.createClientAnchor();
			//i found it useful to show the comment box at the bottom right corner
			anchor.setCol1(cell1.getColumnIndex() + 1); //the box of the comment starts at this given column...
			anchor.setCol2(cell1.getColumnIndex() + 3); //...and ends at that given column
			anchor.setRow1(cell1.getRowIndex() + 1); //one row below the cell...
			anchor.setRow2(cell1.getRowIndex() + 5); //...and 4 rows high

			Drawing drawing = sheet.createDrawingPatriarch();
			Comment comment = drawing.createCellComment(anchor);

			//set the comment text and author
			comment.setString(factory.createRichTextString("Found " + note));
			comment.setAuthor("SYSTEM");

			cell1.setCellComment(comment);
		}
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
		default: throw new Exception(String.format("Unexpected Cell[%s] Type[%s] of Sheet[%s]", cell.getReference(), 
				cell.getCellType(), cell.getSheet().getSheetName()));
		}
		return content;
	}

	public static boolean deleteIfExists(File file) {
		if(file.exists())
			return file.delete();

		return false;		
	}

}
