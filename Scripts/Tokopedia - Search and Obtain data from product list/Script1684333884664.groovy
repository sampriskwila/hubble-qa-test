import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject

import org.openqa.selenium.Keys
import org.openqa.selenium.WebElement

import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI

WebUI.openBrowser('')

WebUI.navigateToUrl('https://www.tokopedia.com/')

WebUI.waitForPageLoad(30)

WebUI.setText(findTestObject('input_search'), 'iphone 14 pro')

WebUI.sendKeys(findTestObject('input_search'), Keys.chord(Keys.ENTER))

WebUI.waitForPageLoad(30)

WebUI.click(findTestObject('official_store_checkbox'))

WebUI.waitForPageLoad(30)

WebUI.scrollToPosition(0, 500)

WebUI.setText(findTestObject('input_price_min'), '100000')

WebUI.sendKeys(findTestObject('input_price_min'), Keys.chord(Keys.ENTER))

WebUI.waitForPageLoad(30)

WebUI.scrollToPosition(0, 500)

WebUI.setText(findTestObject('input_price_max'), '30000000')

WebUI.sendKeys(findTestObject('input_price_max'), Keys.chord(Keys.ENTER))

WebUI.waitForPageLoad(30)

WebUI.click(findTestObject('sort_paling_sesuai'))

WebUI.click(findTestObject('sort_harga_terendah'))

List<WebElement> productPageOne = WebUI.findWebElements(findTestObject('product_item'), 5)

CustomKeywords.'hubblePackage.ExcelFileHandler.writeToExcel'(productPageOne, 'Page 1')

WebUI.scrollToPosition(1000, 9999)

WebUI.verifyElementPresent(findTestObject('page_two'), 10)

WebUI.click(findTestObject('page_two'))

WebUI.waitForPageLoad(30)

List<WebElement> productPageTwo = WebUI.findWebElements(findTestObject('product_item'), 5)

CustomKeywords.'hubblePackage.ExcelFileHandler.writeToExcel'(productPageTwo, 'Page 2')

WebUI.scrollToPosition(1000, 9999)

WebUI.verifyElementPresent(findTestObject('page_three'), 10)

WebUI.click(findTestObject('page_three'))

WebUI.waitForPageLoad(30)

List<WebElement> productPageThree = WebUI.findWebElements(findTestObject('product_item'), 5)

CustomKeywords.'hubblePackage.ExcelFileHandler.writeToExcel'(productPageThree, 'Page 3')

WebUI.closeBrowser()

