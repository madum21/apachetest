package summer.practice.modify;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;

public class ModifyExcel1_1 {

	public static void main(final String[] args) throws Exception {
		final Workbook workbook = new HSSFWorkbook();
		final Sheet sheet = workbook.createSheet("TestHL");
		for (int i = 0; i < 20; i++) {
			final Row row = CellUtil.getRow(i, sheet);
			for (int j = 0; j < 10; j++) {
				final Cell cell = row.createCell(j);
				cell.setCellValue((int) (4 + Math.random() * 6));
			}
		}
		workbook.write(new FileOutputStream(new File("in.xls")));
	}
}
