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
 *  - Create file: excelFile01
 *  - add sheet for file 
 *  - Write some data to File01Sheet0 in File01

 */
//String excelFile01 = 'Data Files\\9.xlsx'
randomstring = CustomKeywords.'defaultpackage.custome.randomString'('letters', 9)

String dirName = System.getProperty('user.dir')

//to print the name on consol and check it 
println('$$$$$' + dirName)
println('$$$$$' + randomstring)

String excelFile01 = ("Data Files//" + randomstring + "xlsx")
//String excelFile01 = randomstring + '.xlsx'

String excelFile2 = randomstring

println('$$$$$' + excelFile01)

// Create file
ExcelKeywords.createExcelFile(excelFile01)

// Verify file is created
File file1 = new File(excelFile01)

assert file1.exists() == true

// Create some new sheets for File01
workbook01 = ExcelKeywords.getWorkbook(excelFile01)

ExcelKeywords.createExcelSheet(workbook01 // Create default sheet 
    )

ExcelKeywords.saveWorkbook(excelFile01, workbook01)

// Write some data to File01Sheet0 in File01
//setValueToCellByAddress(sheet, "A10", 123)
sheet0 = ExcelKeywords.getExcelSheet(workbook01, 'sheet0')

ExcelKeywords.setValueToCellByIndex(sheet0, 0, 0, 'Name')

ExcelKeywords.setValueToCellByIndex(sheet0, 0, 1, 'Legal name')

ExcelKeywords.setValueToCellByIndex(sheet0, 0, 2, 'Common Names')

// Save data of File01
ExcelKeywords.saveWorkbook(excelFile01, workbook01)

//Get name of sheets in workbook
String sheetname = ExcelKeywords.getSheetNames(workbook01)

WebUI.sendKeys(findTestObject('Import Companies/choose to upload file - input file'), dirName +"/"+
excelFile01)


//WebUI.uploadFile(findTestObject('Import Companies/Choose to upload file - input file'), ((dirName + '\\') + excelFile2) + 
   // '.xlsx')

WebUI.click(findTestObject('Import Companies/Import compines button'))

WebUI.click(findTestObject('Import Companies/here link'))

// Delete file


//close browser
WebUI.closeBrowser()
