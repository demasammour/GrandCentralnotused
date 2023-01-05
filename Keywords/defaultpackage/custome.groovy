package defaultpackage

import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject

import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.checkpoint.Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testcase.TestCase
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows

import internal.GlobalVariable

public class custome {
	@Keyword
	public static String randomString(String charsType, int length) {

		String strChar = "abcdefghijklmnopqrstuvwxyz"
		String strNum = "123456789"
		Random rand = new Random()
		StringBuilder sb = new StringBuilder()
		if(charsType=="letters") {
			for (int j=1;j<=length;j++) {
				sb.append(strChar.charAt(rand.nextInt(strChar.length())))
			}
		}
		else if(charsType=="numbers"){
			for (int j=1;j<=length;j++) {
				sb.append(strNum.charAt(rand.nextInt(strNum.length())))
			}
		}


		return sb.toString();
	}
}
