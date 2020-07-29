package summer.practice.modify;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ModifyExcel2 {

	public static void main(final String[] args) throws Exception {
		final Workbook workbook = WorkbookFactory.create(new File("test3.xls"));
		final CreationHelper helper = workbook.getCreationHelper();
		final FormulaEvaluator evaluator = helper.createFormulaEvaluator();
		final int numberOfSheets = workbook.getNumberOfSheets();
		for (int i = 0; i < numberOfSheets; i++) {
			final Sheet sheet = workbook.getSheetAt(i);
			final Drawing drawing = sheet.createDrawingPatriarch();
			System.out.println(sheet.getSheetName());
			for (int rowN = sheet.getFirstRowNum(); rowN <= sheet.getLastRowNum(); rowN++) {
				final Row row = sheet.getRow(rowN);
				for (int col = row.getFirstCellNum(); col < row.getLastCellNum(); col++) {
					final Cell cell = row.getCell(col);
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_FORMULA:
						final String formula = cell.getCellFormula();
						evaluator.evaluateInCell(cell);
						//
						// When the comment box is visible, have it show in a 2x3 space
						final ClientAnchor anchor = helper.createClientAnchor();
						anchor.setCol1(cell.getColumnIndex());
						anchor.setCol2(cell.getColumnIndex() + 2);
						anchor.setRow1(row.getRowNum());
						anchor.setRow2(row.getRowNum() + 3);
						// Create the comment and set the text+author
						final Comment comment = drawing.createCellComment(anchor);
						final RichTextString str = helper.createRichTextString(formula);
						comment.setString(str);
						comment.setAuthor("Summer Practice Student");
						// Assign the comment to the cell
						cell.setCellComment(comment);
						break;
					}
				}
				System.out.println();
			}
		}
		workbook.write(new FileOutputStream(new File("test2_mod.xls")));
	}
}
