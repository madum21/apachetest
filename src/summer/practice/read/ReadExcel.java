package summer.practice.read;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ReadExcel {

	public static void main(final String[] args) throws Exception {
		final InputStream inp = new FileInputStream("test3.xls");
	    final HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(inp));
	    final ExcelExtractor extractor = new ExcelExtractor(wb);

	    extractor.setFormulasNotResults(false);
	    extractor.setIncludeSheetNames(false);
	    final String text = extractor.getText();
	    System.out.println(text);
	    extractor.close();

	}
}
