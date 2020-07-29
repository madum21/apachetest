package summer.practice.read;

import java.io.File;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExcel1 {

	public static void main(final String[] args) throws Exception {
		final Workbook workbook = WorkbookFactory.create(new File("test1.xls"));
		final int numberOfSheets = workbook.getNumberOfSheets();
		for (int i = 0; i < numberOfSheets; i++) {
			final Sheet sheet = workbook.getSheetAt(i);
			System.out.println(sheet.getSheetName());
			for (int rowN = sheet.getFirstRowNum(); rowN <= sheet.getLastRowNum(); rowN++) {
				final Row row = sheet.getRow(rowN);
				for (int col = row.getFirstCellNum(); col < row.getLastCellNum(); col++) {
					final Cell cell = row.getCell(col, Row.RETURN_BLANK_AS_NULL);
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
						System.out.println(cell.getNumericCellValue() + "\t");
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.println(cell.getStringCellValue() + "\t");
						break;
					default:
						break;
					}
					
//					System.out.print(cell.getStringCellValue() + "\t");
				}
				System.out.println();
			}
		}
	}
}
