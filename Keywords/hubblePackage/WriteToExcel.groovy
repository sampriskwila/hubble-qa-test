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

		String filePath = RunConfiguration.getProjectDir() + '/Excel Output File/' + sheetName + '.xlsx'
		FileOutputStream fileOutput = new FileOutputStream(filePath)
		workbook.write(fileOutput)
		fileOutput.close()
	}
}
