package hubblePackage

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.openqa.selenium.WebElement

import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.configuration.RunConfiguration

public class WriteToExcel {
	@Keyword
	public void storeData(List<WebElement> elements, String sheetName) {
		XSSFWorkbook workbook = new XSSFWorkbook()
		XSSFSheet sheet = workbook.createSheet(sheetName)
		
		for (int i = 0; i < elements.size(); i++) {
			Row row = sheet.createRow(i)
			Cell cell = row.createCell(0)
			cell.setCellValue(elements.get(i).getText())
		}
		
		String filePath = RunConfiguration.getProjectDir() + "/Excel Output File/" + sheetName + ".xlsx"
		File file = createFolder(filePath)
		processOutputFile(workbook, file)
	}
	
	private File createFolder(String filePath) {
		File file = new File(filePath)
		file.getParentFile().mkdirs()
		file.createNewFile()
		
		return file
	}
	
	private void processOutputFile(XSSFWorkbook workbook, File file) {
		FileOutputStream fileOutput = new FileOutputStream(file)
		workbook.write(fileOutput)
		fileOutput.close()
	}
}
