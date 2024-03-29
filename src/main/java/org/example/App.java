package org.example;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Hello world!
 */
public class App {
    public static void main(String[] args) throws InterruptedException, IOException {

        System.setProperty("webdriver.chrome.driver", "C:\\proje\\chromedriver\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.get("https://edvgerial.kapitalbank.az/");

        Thread.sleep(3000);
        WebElement mobileNumberInput = driver.findElement(By.id("mobile"));
        mobileNumberInput.sendKeys(" 553349868");

        Thread.sleep(3000);
        WebElement passwordInput = driver.findElement(By.id("password"));
        passwordInput.sendKeys("Parolyaz");

        Thread.sleep(3000);
        WebElement loginButton = driver.findElement(By.xpath("//button[contains(text(),'Daxil ol')]"));
        loginButton.click();

        FileInputStream fis = new FileInputStream("C:\\proje\\Vat\\EDV.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheet("Sheet1");
        int rowCount = sheet.getLastRowNum();
        int colCount = sheet.getRow(0).getLastCellNum();
        System.out.println("rowCount" + rowCount);
        System.out.println("colCount" + colCount);

        for (int i = 0; i <= rowCount; i++) {


            XSSFRow row = sheet.getRow(i);
            XSSFCell cell = row.getCell(0);
            String fiscalId = cell.getStringCellValue();
            Thread.sleep(1105);
            WebElement fiscalIdInput = driver.findElement(By.cssSelector("input.fiscal-id"));
            fiscalIdInput.clear();
            fiscalIdInput.sendKeys(fiscalId);
            Thread.sleep(1550);
            WebElement searchButton = driver.findElement(By.cssSelector("button.submit-receipt"));
            searchButton.click();
            Thread.sleep(2220);

            System.out.println(fiscalId);
            CellStyle style = workbook.createCellStyle();
            style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(style);

        }
        FileOutputStream fos = new FileOutputStream("C:\\proje\\Vat\\EDV.xlsx");
        workbook.write(fos);
        fos.close();


        driver.quit();


    }
}
