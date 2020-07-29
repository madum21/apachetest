package summer.practice.modify;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ModifyExcel1 {

	public static void main(final String[] args) throws Exception {
		final Workbook workbook = WorkbookFactory.create(new File("test3.xls"));
		final FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		final int numberOfSheets = workbook.getNumberOfSheets();
		for (int i = 0; i < numberOfSheets; i++) {
			final Sheet sheet = workbook.getSheetAt(i);
			System.out.println(sheet.getSheetName());
			for (int rowN = sheet.getFirstRowNum(); rowN <= sheet.getLastRowNum(); rowN++) {
				final Row row = sheet.getRow(rowN);
				for (int col = row.getFirstCellNum(); col < row.getLastCellNum(); col++) {
					final Cell cell = row.getCell(col);
					evaluator.evaluateInCell(cell);
				}
				System.out.println();
			}
		}
		workbook.write(new FileOutputStream(new File("test1_mod.xls")));
	}
}
