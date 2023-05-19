package hubblePackage;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebElement;

import com.kms.katalon.core.annotation.Keyword;
import com.kms.katalon.core.configuration.RunConfiguration;

public class ExcelFileHandler {
	@Keyword
	public void writeToExcel(List<WebElement> elements, String sheetName) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet(sheetName);
		
		for (int i = 0; i < elements.size(); i++) {
			sheet.createRow(i).createCell(0).setCellValue(elements.get(i).getText());
		}

		String filePath = RunConfiguration.getProjectDir() + "/Excel Output File/" + sheetName + ".xlsx";
		File file = createFolder(filePath);
		processOutputFile(workbook, file);
	}

	private File createFolder(String filePath) throws IOException {
		File file = new File(filePath);
		file.getParentFile().mkdirs();
		file.createNewFile();

		return file;
	}

	private void processOutputFile(XSSFWorkbook workbook, File file) throws IOException {
		FileOutputStream fileOutput = new FileOutputStream(file);
		workbook.write(fileOutput);
		fileOutput.close();
	}
}
