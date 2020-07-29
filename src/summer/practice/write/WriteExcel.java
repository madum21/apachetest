package summer.practice.write;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class WriteExcel {

	public static void main(final String[] args) throws Exception {
		final Workbook workbook = new HSSFWorkbook();
		final Sheet sheet = workbook.createSheet("SummerPractice");
		final Row row = sheet.createRow(0);
		Cell cell = row.createCell(0, Cell.CELL_TYPE_NUMERIC);
		cell.setCellValue(2);
		cell = row.createCell(1, Cell.CELL_TYPE_NUMERIC);
		cell.setCellValue(5);
		cell = row.createCell(2, Cell.CELL_TYPE_FORMULA);
		cell.setCellFormula("SUM(A1:B1)");
		//
		final CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setBorderLeft(CellStyle.BORDER_THICK);
		cellStyle.setBorderRight(CellStyle.BORDER_DOUBLE);
		cellStyle.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
		cellStyle.setFillPattern(CellStyle.THICK_BACKWARD_DIAG);
		cellStyle.setFillBackgroundColor(HSSFColor.YELLOW.index);
		final Font font = workbook.createFont();
		font.setColor(HSSFColor.RED.index);
		cellStyle.setFont(font);
		cell.setCellStyle(cellStyle);
		//
		workbook.write(new FileOutputStream(new File("test_out.xls")));
		System.out.println(sheet.getSheetName());
	}
}
