package summer.practice.read;

import java.io.File;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExcel4 {

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
					final CellValue cellValue = evaluator.evaluate(cell);
					switch (cellValue.getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cellValue.getNumberValue() + "\t");
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.print(cellValue.getStringValue() + "\t");
						break;
					case Cell.CELL_TYPE_FORMULA:
						// not possible
						break;
					}
				}
				System.out.println();
			}
		}
	}
}
