package summer.practice.read;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.FontFormatting;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Read {

	public static void main(final String[] args) throws Exception {
		final Workbook workbook = WorkbookFactory.create(new File("in.xls"));
		final Sheet sheet = workbook.getSheetAt(0);
		for (int rowN = sheet.getFirstRowNum(); rowN <= sheet.getLastRowNum(); rowN++) {
			final Row row = sheet.getRow(rowN);
			final int r1 = rowN + 1;
			final String col = col(row.getLastCellNum() - 1);
			final Cell createCell = row.createCell(row.getLastCellNum());
			createCell.setCellFormula("AVERAGE(" + col(0) + r1 + ":" + col + r1 + ")");
		}

		workbook.write(new FileOutputStream(new File("out.xls")));
	}

	private static ConditionalFormattingRule c1(final Sheet sheet) {
		final SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

		final ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule(ComparisonOperator.LT, "5");
		final FontFormatting fontFmt = rule1.createFontFormatting();
		fontFmt.setFontStyle(true, false);
		fontFmt.setFontColorIndex(IndexedColors.RED.index);

		return rule1;
	}

	private static ConditionalFormattingRule c1(final Sheet sheet, final String c) {
		final SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();
		final ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.GT, c);
		final FontFormatting fontFmt2 = rule2.createFontFormatting();
		fontFmt2.setFontStyle(false, true);
		fontFmt2.setFontColorIndex(IndexedColors.GREEN.index);
		return rule2;
	}

	private static String col(final int col) {
		return "" + (char) ('A' + col);
	}
}
