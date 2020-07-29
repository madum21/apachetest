package summer.practice.modify;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.BorderFormatting;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.FontFormatting;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

public class ModifyExcel {

	public static void main(final String[] args) throws Exception {
		final Workbook workbook = WorkbookFactory.create(new File("test3.xls"));
		final Sheet sheet = workbook.getSheetAt(0);
		//
		final SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

		final ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule(ComparisonOperator.LT, "F1");
		final FontFormatting fontFmt = rule1.createFontFormatting();
		fontFmt.setFontStyle(true, false);
		fontFmt.setFontColorIndex(IndexedColors.DARK_RED.index);

		final BorderFormatting bordFmt = rule1.createBorderFormatting();
		bordFmt.setBorderBottom(BorderFormatting.BORDER_THIN);
		bordFmt.setBorderTop(BorderFormatting.BORDER_THICK);
		bordFmt.setBorderLeft(BorderFormatting.BORDER_DASHED);
		bordFmt.setBorderRight(BorderFormatting.BORDER_DOTTED);

		final PatternFormatting patternFmt = rule1.createPatternFormatting();
		patternFmt.setFillBackgroundColor(IndexedColors.YELLOW.index);

		final ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.BETWEEN, "2", "3");
		final FontFormatting fontFmt2 = rule2.createFontFormatting();
		fontFmt2.setFontStyle(false, true);
		fontFmt2.setFontColorIndex(IndexedColors.GREEN.index);

		final ConditionalFormattingRule[] cfRules = { rule1, rule2 };

		final CellRangeAddress[] regions = { CellRangeAddress.valueOf("A1:E1") };

		sheetCF.addConditionalFormatting(regions, cfRules);
		//
		workbook.write(new FileOutputStream(new File("test_mod.xls")));
	}
}
