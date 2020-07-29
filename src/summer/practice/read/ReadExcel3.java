package summer.practice.read;

import java.io.File;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExcel3 {

	public static void main(final String[] args) throws Exception {
		final Workbook workbook = WorkbookFactory.create(new File("test3.xls"));
		final int numberOfSheets = workbook.getNumberOfSheets();
		for (int i = 0; i < numberOfSheets; i++) {
			final Sheet sheet = workbook.getSheetAt(i);
			System.out.println(sheet.getSheetName());
			for (int rowN = sheet.getFirstRowNum(); rowN <= sheet.getLastRowNum(); rowN++) {
				final Row row = sheet.getRow(rowN);
				for (int col = row.getFirstCellNum(); col < row.getLastCellNum(); col++) {
					final Cell cell = row.getCell(col);
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cell.getNumericCellValue() + "\t");
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.print(cell.getStringCellValue() + "\t");
						break;
					case Cell.CELL_TYPE_FORMULA:
						System.out.print(cell.getCellFormula() + "\t");
						break;
					}
				}
				System.out.println();
			}
		}
	}
}
