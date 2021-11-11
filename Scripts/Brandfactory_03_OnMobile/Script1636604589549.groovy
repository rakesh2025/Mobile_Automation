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
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.*;

Mobile.startApplication('C:\\Users\\DeLL\\Downloads\\Brand_Factory_base.apk', false)
Mobile.tap(findTestObject('Object Repository/BrandFactory_Mobile/android.widget.ImageView'), 0)
Mobile.tap(findTestObject('Object Repository/BrandFactory_Mobile/android.widget.ImageView (1)'), 0)
Mobile.tap(findTestObject('Object Repository/BrandFactory_Mobile/android.widget.ImageView (2)'), 0)
Mobile.tap(findTestObject('Object Repository/BrandFactory_Mobile/android.widget.TextView - Jeans'), 0)
Mobile.tap(findTestObject('Object Repository/BrandFactory_Mobile/android.widget.ImageView (3)'), 0)
Mobile.scrollToText('32')
Mobile.tap(findTestObject('Object Repository/BrandFactory_Mobile/android.widget.TextView - 32'), 0)
Mobile.tap(findTestObject('Object Repository/BrandFactory_Mobile/android.widget.TextView - ADD TO BAG'), 0)
Mobile.tap(findTestObject('Object Repository/BrandFactory_Mobile/android.widget.ImageView (4)'), 0)
def textOnMobile=Mobile.getText(findTestObject('Object Repository/BrandFactory_Mobile/android.widget.TextView - LAWMANPg3'), 0)
writeInExcel(textOnMobile)
Mobile.tap(findTestObject('Object Repository/BrandFactory_Mobile/android.widget.TextView - PROCEED TO CHECKOUT'), 0)
Mobile.tap(findTestObject('Object Repository/BrandFactory_Mobile/android.widget.TextView - CONFIRM  PAY'), 0)
Mobile.closeApplication()

def writeInExcel(def data) {
	FileInputStream fis=new FileInputStream("C:\\Users\\DeLL\\Desktop\\data.xlsx");
	Workbook wb= WorkbookFactory.create(fis);
	Sheet sh=wb.getSheet("Sheet1");
	Row row=sh.getRow(1);
	Cell cell=row.createCell(3);
	cell.setCellValue(data);
	FileOutputStream fos=new FileOutputStream("C:\\Users\\DeLL\\Desktop\\data.xlsx");
	wb.write(fos);
	wb.close();
	
   }
