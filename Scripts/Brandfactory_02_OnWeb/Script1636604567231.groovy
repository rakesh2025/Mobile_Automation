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
import java.io.FileInputStream as FileInputStream
import org.apache.poi.ss.usermodel.*

WebUI.openBrowser('')
WebUI.navigateToUrl('https://www.brandfactoryonline.com/')
WebUI.mouseOver(findTestObject('BrandFactory_03_Web/Page_Brand Factory Online  Online Shopping for Clothes  Fashion Accessories in India/Account_img'))
WebUI.click(findTestObject('Object Repository/BrandFactory_03_Web/Page_Brand Factory Online  Online Shopping _a889ab/a_LOG IN'))
WebUI.setText(findTestObject('Object Repository/BrandFactory_03_Web/Page_BrandFactory Login Customer/input_Using Email  Mobile Number_login-email'), 
    '7008145134')
WebUI.setEncryptedText(findTestObject('Object Repository/BrandFactory_03_Web/Page_BrandFactory Login Customer/input_Using Email  Mobile Number_login-pswd'), 
    'Hobzp+enCtI=')
WebUI.click(findTestObject('Object Repository/BrandFactory_03_Web/Page_BrandFactory Login Customer/button_LOG IN'))
WebUI.click(findTestObject('BrandFactory_03_Web/Page_Brand Factory Online  Online Shopping for Clothes  Fashion Accessories in India/Cart_Img'))
def nameOnWeb = WebUI.getText(findTestObject('BrandFactory_03_Web/Page_BrandFactory/Captured_Text_From_Web'))
println nameOnWeb+"-----------------------------------------------------"
println readFromExcel(1, 3)
WebUI.verifyEqual(nameOnWeb, readFromExcel(1, 3))
WebUI.closeBrowser()

def readFromExcel(def rowno, def cellno) {
    FileInputStream fis = new FileInputStream('C:\\Users\\DeLL\\Desktop\\data.xlsx')
    Workbook wb = WorkbookFactory.create(fis)
    Sheet sh = wb.getSheet('Sheet1')
    Row row = sh.getRow(rowno)
    Cell cell = row.getCell(cellno)
    def data = cell.getStringCellValue()
    wb.close()
    return data
}

