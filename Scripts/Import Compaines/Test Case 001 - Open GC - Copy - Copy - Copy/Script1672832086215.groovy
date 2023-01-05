import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys
import com.kms.katalon.keyword.excel.ExcelKeywords as ExcelKeywords

WebUI.callTestCase(findTestCase('Login/Test Case 001 - Login'), [:], FailureHandling.STOP_ON_FAILURE)

WebUI.click(findTestObject('Main Menu/Contacts main menu'))

WebUI.click(findTestObject('Main Menu/Import companies'))

/**
 * Example for interact with excel using custom keywords
 *  Scenarios:
 *  - Create two files: File01.xlsx, File02.xlsx
 *  - Add some content to file File01
 *  - Read data from file File01, then set these values to File02 
 *  - Compare File01 with File02 (cell, row, sheet, file)
 *  - Clean up data (delete all files)
 */
String excelFile01 = 'Data Files\\File01.xlsx'

// Create two files
ExcelKeywords.createExcelFile(excelFile01)


// Create some new sheets for File01
workbook01 = ExcelKeywords.getWorkbook(excelFile01)

ExcelKeywords.createExcelSheet(workbook01) // Create default sheet 
ExcelKeywords.createExcelSheets(workbook01, ['File01Sheet01', 'File01Sheet02']) // Create list of sheets with name
ExcelKeywords.saveWorkbook(excelFile01, workbook01)

// Create some new sheets for File01
workbook01 = ExcelKeywords.getWorkbook(excelFile01)

ExcelKeywords.createExcelSheet(workbook01) // Create default sheet
	

ExcelKeywords.saveWorkbook(excelFile01, workbook01)

// Write some data to File01Sheet0 in File01
//setValueToCellByAddress(sheet, "A10", 123)
sheet0 = ExcelKeywords.getExcelSheet(workbook01, 'sheet0')

ExcelKeywords.setValueToCellByIndex(sheet0, 0, 0, 'Name')

ExcelKeywords.setValueToCellByIndex(sheet0, 0, 1, 'Legal name')

ExcelKeywords.setValueToCellByIndex(sheet0, 0, 2, 'Common Names')


// Save data of File01
ExcelKeywords.saveWorkbook(excelFile01, workbook01)
WebUI.sendKeys(findTestObject('Import Companies/choose to upload file - input file'), '/Users/dimaabusamour/git/GrandCentral/Data Files/Data Files/File01.xlsx')
WebUI.click(findTestObject('Import Companies/Import compines button'))

