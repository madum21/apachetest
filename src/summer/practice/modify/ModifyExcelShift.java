package summer.practice.modify;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellUtil;

public class ModifyExcelShift {

	public static void main(final String[] args) throws Exception {
		final Workbook workbook = WorkbookFactory.create(new File("Book2.xlsx"));
		final Sheet sheet = workbook.getSheetAt(6);
		sheet.shiftRows(9, 14, 1);
		final Row row = CellUtil.getRow(9, sheet);
		final Cell cell = CellUtil.getCell(row, 0);
		cell.setCellValue(11111);
		workbook.write(new FileOutputStream(new File("Book1_out.xlsx")));
	}
}
