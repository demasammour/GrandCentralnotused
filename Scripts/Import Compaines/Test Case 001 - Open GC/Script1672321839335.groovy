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

WebUI.sendKeys(findTestObject('Import Companies/choose to upload file - input file'), '/Users/dimaabusamour/git/files for importing compines/first one.xlsx')

WebUI.click(findTestObject('Import Companies/Import compines button'))

WebUI.click(findTestObject('Import Companies/here link'))

//String excelFilePath =  '/Users/dimaabusamour/git/GrandCentral/Data Files/Importing compines files/first one.xlsx'
String excelFile01 = '//Data Files//Importing compines files//GC - Run 1 copy.xlsx'

// Create file
ExcelKeywords.createExcelFile(excelFile01)

// Verify file created
File file1 = new File(excelFile01)

assert file1.exists() == true

// Create new sheets for File01
workbook01 = ExcelKeywords.getWorkbook(excelFile01)

ExcelKeywords.createExcelSheet(workbook01)// Create default sheet 
    

ExcelKeywords.saveWorkbook(excelFile01, workbook01)

// Verify sheets created
//String[ ] ExpectedListSheetFile1 = ['Sheet0','Sheet1','File01Sheet01', 'File01Sheet02']
String[] ExpectedListSheetFile1 = ['Sheet0']

workbookFile01 = ExcelKeywords.getWorkbook(excelFile01) // get latest workbook File01
    

assert ExcelKeywords.getSheetNames(workbookFile01) == ExpectedListSheetFile1

// Write some data to File01Sheet01 in File01
workbook01 = ExcelKeywords.getWorkbook(excelFile01)

//should be deleted //sheet0 = ExcelKeywords.getExcelSheet(workbook01, 'File01Sheet01')
sheet0 = ExcelKeywords.getExcelSheet(workbook01)

ExcelKeywords.setValueToCellByIndex(sheet0, 0, 0, 'Fruites')

ExcelKeywords.setValueToCellByIndex(sheet0, 1, 0, 'Apple')

ExcelKeywords.setValueToCellByIndex(sheet0, 2, 0, 'Orange')

ExcelKeywords.setValueToCellByAddress(sheet0, 'B1', 'Price')

ExcelKeywords.setValueToCellByAddress(sheet0, 'B2', 10000)

ExcelKeywords.setValueToCellByAddress(sheet0, 'B3', 15000)

ExcelKeywords.setValueToCellByAddress(sheet0, 'C1', 'Quantity')

ExcelKeywords.setValueToCellByAddress(sheet0, 'C2', 5)

ExcelKeywords.setValueToCellByAddress(sheet0, 'C3', 6)

ExcelKeywords.setValueToCellByAddress(sheet0, 'D1', 'Total Prices')

ExcelKeywords.setValueToCellByAddress(sheet0, 'D2', '=B2*C2')

ExcelKeywords.setValueToCellByAddress(sheet0, 'D3', '=B3*C3')

ExcelKeywords.setValueToCellByAddress(sheet0, 'E1', 'Booking date')

ExcelKeywords.setValueToCellByAddress(sheet0, 'E2', '02/28/2019')

ExcelKeywords.setValueToCellByAddress(sheet0, 'E3', '02/29/2019')

// Save data of File01
ExcelKeywords.saveWorkbook(excelFile01, workbook01)
workbook01 = ExcelKeywords.getWorkbook(excelFile01)

//Get name of sheets in workbook
ExcelKeywords.getSheetNames(workbook01)

