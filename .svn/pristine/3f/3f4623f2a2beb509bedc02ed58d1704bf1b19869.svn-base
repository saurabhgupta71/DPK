package com.otc.selenium;

import java.awt.Robot;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Set;

import javax.swing.JOptionPane;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class GenericMethods {

    Driver drivers = new Driver();

    // Declaring all the variables used.
    int inc = 0;
    int summarycounter = 1;
    static String index, errorsnapshotforpass, errorsnapshotforfail, browser, resultdirectory, iedriver, chromedriver, error, errortype, config, htmldriver;
    HSSFWorkbook exeresultworkbook;
    static HSSFSheet exeresultsheet;
    static HSSFRow exerow;
    static HSSFCell execell;
    FileOutputStream resultexe;
    String executionstatus = null;
    org.openqa.selenium.Capabilities cap;
    static String errorsnappath;
    int k;
    public Robot robot;
    String url;
    public Driver d;
    int rowvalue;
    String StrGetText;
    String StrGetValue;
    String result;
    private static String[] links = null;
    private static int linksCount = 0;
    ArrayList<String> obj = new ArrayList<String>();
    String executionresultpath;
    static String credential = null;
    static String skuid = null;
    String text;
    String mainwindow = null;
    String childwindow = null;
    int rowcount = 0;
    String sheetname = null;
    int sheetindex;
    Boolean errorcheck = false;
    String activesheet = null;

    int clearcount = 0;
    int totalrow;
    String[][] innercell = null;
    int rcount = 0;

    public static WebDriver driver;

    public GenericMethods(Driver d) {
        this.d = d;
    }

    // Opening the browser and maximizing it, FF, CHROME & IE.
    public String openFirefoxBrowser(String pStrValue, int i) throws InterruptedException, IOException {
        driver = new FirefoxDriver();
        driver.manage().window().maximize();
        // cap = ((RemoteWebDriver) driver).getCapabilities();
        return pStrValue;
    }

    public void openChomeBrowser(String pStrValue, int i) throws IOException {

        System.setProperty("webdriver.chrome.driver", chromedriver);
        driver = new ChromeDriver();
        driver.manage().window().maximize();
        cap = ((RemoteWebDriver) driver).getCapabilities();

    }

    public void openIEBrowser(String pStrValue) throws IOException {
        File IEDriver = new File(iedriver);
        System.setProperty("webdriver.ie.driver", IEDriver.getAbsolutePath());
        driver = new InternetExplorerDriver();
        cap = ((RemoteWebDriver) driver).getCapabilities();

    }

    // Selecting the browser type FF, CHROME & IE.
    public void browserType(String pStrValue) throws InterruptedException, IOException, Exception {

        if (pStrValue.equalsIgnoreCase("Firefox")) {
            openFirefoxBrowser(pStrValue, i);
        } else if (pStrValue.equalsIgnoreCase("Chrome")) {
            openChomeBrowser(pStrValue, i);
        } else if (pStrValue.equalsIgnoreCase("IE")) {
            openIEBrowser(pStrValue);
        } else {
            System.out.println("PLEASE ENTER Firefox or Chrome or IE NAME");
        }
    }

    // ---- Various methods which are called from the switch cases mentioned below of the page. ----//

    // To hit the URL
    public String openbrowser(String pStrValue) throws InterruptedException, IOException, Exception {
        driver.get(pStrValue);
        return pStrValue;
    }

    // To close the browser
    public void closebrowser() throws IOException, Exception {
        driver.close();
        driver.quit();

    }

    public void ClearBrowserCache() throws InterruptedException {
        driver.manage().deleteAllCookies(); // delete all cookies
        Thread.sleep(5000); // wait 5 seconds to clear cookies.
    }

    // To close the promotional pop-up coming in the home page of the OTC application
    public void promotionalPopupCheck(String pStrObjectDescription) {
        try {
            if (driver.findElement(By.id("ltbx_init")).isDisplayed()) {
                driver.findElement(By.xpath(".//*[@id='ltbx_initMap']/area[4]")).click();

                Set<String> windows = driver.getWindowHandles();
                for (String window : windows) {
                    driver.switchTo().window(window);
                }
            }
        } catch (Exception e) {
            System.out.println(e);
            e.printStackTrace();
        }

    }

    public void popUpHandle(String pStrObjectDescription) {
        // mainwindow = driver.getWindowHandle();

        try {
            Set<String> handles = driver.getWindowHandles();// To handle multiple windows
            mainwindow = driver.getWindowHandle(); // To get your main window
            handles.remove(mainwindow);
            String winHandle = handles.iterator().next(); // To find popup window
            if (winHandle != mainwindow) {
                childwindow = winHandle;

                driver.switchTo().window(childwindow);
            }
        } catch (Exception e) {
            System.out.println(e);
            e.printStackTrace();
        }

    }

    String mainwindow1;

    public void defaultWindow() {
        try {
            mainwindow1 = driver.getWindowHandle();

            driver.switchTo().window(mainwindow);
        } catch (Exception e) {
            System.out.println(e);
        }
    }

    public void popUpHandleClose() {
        try {

            driver.close();
            driver.switchTo().window(mainwindow);
        } catch (Exception e) {

            System.out.println(e);
            e.printStackTrace();
        }
    }

    // Separate method to identify which locator is called.
    public static By put(String pStrObjectName, String pStrObjectDescription) throws Exception {
        if (pStrObjectName.equalsIgnoreCase("id")) {
            return By.id(pStrObjectDescription);
        } else if (pStrObjectName.equalsIgnoreCase("name")) {
            return By.name(pStrObjectDescription);
        } else if (pStrObjectName.equalsIgnoreCase("className")) {
            return By.className(pStrObjectDescription);
        } else if (pStrObjectName.equalsIgnoreCase("linktext")) {
            return By.linkText(pStrObjectDescription);
        } else {
            return By.xpath(pStrObjectDescription);
        }
    }

    // Explicit wait for element to appear.
    public static void waitUntilElementIsPresent(String pStrObjectName, String pStrObjectDescription) throws Exception {
        try {

            WebDriverWait wait = new WebDriverWait(driver, 30);
            wait.until(ExpectedConditions.visibilityOfElementLocated(put(pStrObjectName, pStrObjectDescription)));
            // wait.until(ExpectedConditions.elementToBeClickable(put(pStrObjectName, pStrObjectDescription)));
            wait.until(ExpectedConditions.presenceOfElementLocated(put(pStrObjectName, pStrObjectDescription)));

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    // Clear the value from TB
    public void TextBoxClear(String pStrObjectName, String pStrObjectDescription) throws Exception {
        waitUntilElementIsPresent(pStrObjectName, pStrObjectDescription);
        driver.findElement(put(pStrObjectName, pStrObjectDescription)).clear();
    }

    // Type the value in TB
    public void TextBoxType(String pStrObjectName, String pStrObjectDescription, String pStrvalue) throws IOException, Exception {
        waitUntilElementIsPresent(pStrObjectName, pStrObjectDescription);
        driver.findElement(put(pStrObjectName, pStrObjectDescription)).sendKeys(pStrvalue);

    }

    // To click on any element
    public void Click(String pStrObjectName, String pStrObjectDescription) throws InterruptedException, IOException, Exception {
        waitUntilElementIsPresent(pStrObjectName, pStrObjectDescription);
        driver.findElement(put(pStrObjectName, pStrObjectDescription)).click();
    }

    public void Submit(String pStrObjectName, String pStrObjectDescription) throws InterruptedException, IOException, Exception {
        waitUntilElementIsPresent(pStrObjectName, pStrObjectDescription);
        driver.findElement(put(pStrObjectName, pStrObjectDescription)).submit();
    }

    // To check visibility of any element
    public void IsVisible(String pStrObjectName, String pStrObjectDescription) throws InterruptedException, IOException, Exception {
        waitUntilElementIsPresent(pStrObjectName, pStrObjectDescription);
        driver.findElement(put(pStrObjectName, pStrObjectDescription)).isDisplayed();
    }

    // To verify any text present in the page.
    public void IsTextPresent(String pStrObjectName, String pStrObjectDescription, String pStrValue) throws IOException, InterruptedException, Exception {
        waitUntilElementIsPresent(pStrObjectName, pStrObjectDescription);
        text = pStrValue.trim();
        if (driver.findElement(put(pStrObjectName, pStrObjectDescription)).getText().equalsIgnoreCase(text)) {

        } else {
            throw new Exception("Expected Value \"" + text + "\" not match with Actual Value \"" + driver.findElement(put(pStrObjectName, pStrObjectDescription)).getText() + "\"");
        }
    }

    // To check if Radio Button or Check Box is selected or not
    public void IsSelected(String pStrObjectName, String pStrObjectDescription) throws IOException, Exception {
        driver.findElement(put(pStrObjectName, pStrObjectDescription)).isSelected();
        if (driver.findElement(put(pStrObjectName, pStrObjectDescription)).isSelected() == true) {

        } else {
            throw new Exception("Passed value not selected");
        }
    }

    // To get the data from textbox
    public void getText(String pStrObjectName, String pStrObjectDescription) throws Exception {
        result = driver.findElement(put(pStrObjectName, pStrObjectDescription)).getText();
    }

    // Refresh the page
    public void Refresh() {
        driver.navigate().refresh();
    }

    // To navigate into the new URL
    public void To(String pStrObjectDescription) {
        driver.navigate().to(pStrObjectDescription);

    }

    // To navigate page forward
    public void Forward() {
        driver.navigate().forward();
    }

    // To navigate page backward
    public void Backward() {
        driver.navigate().back();
    }

    // To get the current url
    public String GetCurrentUrl(WebDriver driver) {
        url = driver.getCurrentUrl();

        return url;
    }

    public void switchToiFrame(String iframe) {
        try {
            driver.switchTo().frame(iframe);
        } catch (Exception e) {
            System.out.println(e);
            e.printStackTrace();
        }
    }

    public void switchToDefault() throws Exception {

        driver.switchTo().defaultContent();

    }

    public void iFrameTextBoxType(String pStrObjectName, String pStrObjectDescription, String pStrValue) throws Exception {
        WebDriverWait wait = new WebDriverWait(driver, 30);

        wait.until(ExpectedConditions.presenceOfElementLocated(By.tagName("iframe")));
        int totalframe = driver.findElements(By.tagName("iframe")).size();
        // List<WebElement> element = driver.findElements(By.tagName("iframe"));

        // System.out.println(element);
        for (int iframec = 0; iframec <= totalframe;) {
            // driver.switchTo().frame(0);
            driver.switchTo().frame(driver.findElement(By.tagName("iframe")));
            int count = driver.findElements(put(pStrObjectName, pStrObjectDescription)).size();

            // int count = driver.findElements(By.xpath(//div[@class='email-signup)).size();

            if (count > 0) {
                iframec++;
                driver.findElement(put(pStrObjectName, pStrObjectDescription)).sendKeys(pStrValue);
                break;
            } else {
                iframec++;
                continue;

            }

        }

        driver.switchTo().defaultContent();
    }

    public void iFrameButtonClick(String pStrObjectName, String pStrObjectDescription) throws Exception {
        WebDriverWait wait = new WebDriverWait(driver, 30);

        wait.until(ExpectedConditions.presenceOfElementLocated(By.tagName("iframe")));
        int totalframe = driver.findElements(By.tagName("iframe")).size();
        // List<WebElement> element = driver.findElements(By.tagName("iframe"));

        // System.out.println(element);
        for (int iframecount = 0; iframecount <= totalframe; iframecount++) {
            // driver.switchTo().frame(i);
            driver.switchTo().frame(driver.findElement(By.tagName("iframe")));
            int count = driver.findElements(put(pStrObjectName, pStrObjectDescription)).size();
            // int count = driver.findElements(By.xpath(//div[@class='email-signup)).size();

            if (count == 1) {

                driver.findElement(put(pStrObjectName, pStrObjectDescription)).click();
                break;
            } else {
                System.out.println("continue looping");
                // driver.switchTo().defaultContent();
            }
        }

        driver.switchTo().defaultContent();

    }

    public void switchToiFrame(WebElement iframe) throws Exception {
        try {
            driver.switchTo().frame(iframe);
        } catch (Exception e) {
            System.out.println(e);
            e.printStackTrace();
        }

    }

    public void SelectByValue(String pStrObjectName, String pStrObjectDescription, String pStrValue) throws Exception {

        Select selectByValue = new Select(driver.findElement(put(pStrObjectName, pStrObjectDescription)));
        selectByValue.selectByValue(pStrValue);
    }

    public void SelectByVisibleText(String pStrObjectName, String pStrObjectDescription, String pStrValue) throws Exception {
        Select selectByVisibleText = new Select(driver.findElement(put(pStrObjectName, pStrObjectDescription)));
        selectByVisibleText.selectByVisibleText(pStrValue);
    }

    public void SelectByIndex(String pStrObjectName, String pStrObjectDescription, String pStrValue) throws Exception {
        Select selectByIndex = new Select(driver.findElement(put(pStrObjectName, pStrObjectDescription)));
        int value = Integer.parseInt(pStrValue);
        selectByIndex.selectByIndex(value);
    }

    // Mouse Hover
    public void mouseHover(String pStrObjectName, String pStrObjectDescription) throws Exception {
        waitUntilElementIsPresent(pStrObjectName, pStrObjectDescription);
        Actions action = new Actions(driver);
        WebElement we = driver.findElement(put(pStrObjectName, pStrObjectDescription));
        action.moveToElement(we).build().perform();
    }

    public void alertHandle() {
        Alert simpleAlert = driver.switchTo().alert();
        String alertText = simpleAlert.getText();
        System.out.println("Alert text is " + alertText);
        simpleAlert.accept();
    }

    // Initializing the variables used.
    int i;
    int f;
    HSSFWorkbook ObjWorkBook;
    HSSFSheet sheet;
    HSSFRow innerRowCount;
    Date date;
    int counter;
    String[][] innerarrayNew = null;
    HSSFRow innerRowCountNew;

    // Writing the test action as Pass in the summary sheet
    public void WritePass(int rowval, String pStrStepID, String TestCase_Type, String pStrTestCaseName, String filepath, int rowvalue) {

        try {

            HSSFCellStyle style = exeresultworkbook.createCellStyle();
            style.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
            style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            if (pStrStepID != null && TestCase_Type != null && pStrTestCaseName != null) {
                exerow = exeresultsheet.createRow(rowval);
                execell = exerow.createCell(0);
                execell.setCellValue(pStrStepID);
                execell = exerow.createCell(1);
                execell.setCellValue(TestCase_Type);
                execell = exerow.createCell(2);
                execell.setCellValue(pStrTestCaseName);

                execell = exerow.createCell(3);
                execell.setCellValue(datetime("dd-MM-yy HH:mm:ss"));
                execell = exerow.createCell(4);
                execell.setCellStyle(style);
                execell.setCellValue("Pass");

                if (rowval == rowvalue) {

                    execell = exerow.createCell(5);
                    execell.setCellValue(result);

                }

                execell = exerow.createCell(6);
                if (errorsnapshotforpass.equalsIgnoreCase("yes")) {
                    execell.setCellValue(snapshotcapture(filepath));
                } else {
                    execell.setCellValue("");
                }

                execell = exerow.createCell(7);
                execell.setCellValue("");
            }
        } catch (Exception e) {
            System.out.println("error in write pass block" + e);
            e.printStackTrace();
        }

    }

    // Writing the test action as Fail in the summary sheet
    public void WriteFail(int rowval, String pStrStepID, String TestCase_Type, String pStrTestCaseName, String execp, String filepath, boolean exeval) {

        try {
            HSSFCellStyle style = exeresultworkbook.createCellStyle();
            style.setFillForegroundColor(HSSFColor.ROSE.index);
            style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            exerow = exeresultsheet.createRow(rowval);
            execell = exerow.createCell(0);
            execell.setCellValue(pStrStepID);
            execell = exerow.createCell(1);
            execell.setCellValue(TestCase_Type);

            execell = exerow.createCell(2);
            execell.setCellValue(pStrTestCaseName);

            execell = exerow.createCell(3);
            execell.setCellValue(datetime("dd-MM-yy HH:mm:ss"));
            if (exeval) {
                execell = exerow.createCell(4);
                execell.setCellStyle(style);
                execell.setCellValue("Not Executed");
            } else {
                execell = exerow.createCell(4);
                execell.setCellStyle(style);
                execell.setCellValue("Fail");
            }

            execell = exerow.createCell(5);
            execell.setCellValue(execp);
            execell = exerow.createCell(6);
            if (errorsnapshotforfail.equalsIgnoreCase("yes")) {
                execell.setCellValue(snapshotcapture(filepath));
            } else {
                execell.setCellValue("");
            }
        } catch (Exception e) {
            System.out.println("error in write fail block" + e);
            e.printStackTrace();
        }

    }

    public void WriteNotExecuted(int rowval, String pStrStepID, String TestCase_Type, String pStrTestCaseName, boolean exeval) {

        HSSFCellStyle style = exeresultworkbook.createCellStyle();
        style.setFillForegroundColor(HSSFColor.LIGHT_ORANGE.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        Integer rowcount = rowval;
        System.out.println("row value" + rowval);

        try {

            if (pStrStepID != null && TestCase_Type != null && pStrTestCaseName != null) {

                exerow = exeresultsheet.createRow(rowval);
                execell = exerow.createCell(0);
                execell.setCellValue(pStrStepID);

                execell = exerow.createCell(1);

                execell.setCellValue(TestCase_Type);

                execell = exerow.createCell(2);
                execell.setCellValue(pStrTestCaseName);

                execell = exerow.createCell(3);
                execell.setCellValue(datetime("dd-MM-yy HH:mm:ss"));
                if (exeval) {
                    execell = exerow.createCell(4);
                    execell.setCellStyle(style);
                    execell.setCellValue("Not Executed");
                } else {
                    execell = exerow.createCell(4);
                    execell.setCellStyle(style);
                    execell.setCellValue("Fail");
                }

            }

        } catch (Exception e) {
            System.out.println("error in write not executed block" + e);
            // e.printStackTrace();

        }

    }

    // Capturing the Snapshot
    public String snapshotcapture(String filepath) throws IOException, Exception {
        String path = snapshotpath(filepath);
        try {

            File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
            org.apache.commons.io.FileUtils.copyFile(scrFile, new File(path));

        } catch (Exception e) {
            e.getMessage();
            e.printStackTrace();
        }
        return path;
    }

    // Returning path of the snapshot captured along with date and in the image format
    public String snapshotpath(String filepath) {
        Date currentdate = new Date();
        SimpleDateFormat dateformat = new SimpleDateFormat("HHmmss_ddMMyyyy");
        return filepath + dateformat.format(currentdate) + ".jpg";
    }

    // Returning the date and time
    public String datetime(String format) {
        Date currentdate = new Date();
        SimpleDateFormat dateformat = new SimpleDateFormat(format);
        return dateformat.format(currentdate);
    }

    // In ExcelRead method all the logic for making the directory for Summary File creation, sheet's iteration is defined, here all logic for excel read is performed.
    public void ExcelRead(File[] InputExcelPath) {

        try {

            String[][] innerarray = null;
            File[] whitelist = new File[InputExcelPath.length];
            if (whitelist != null) {
                for (int i = 0; i < InputExcelPath.length; i++) {
                    String path = InputExcelPath[i].getPath();
                    if (path.contains("Config")) {

                        config(InputExcelPath[i]);

                    } else {

                        whitelist[i] = InputExcelPath[i];

                    }
                }
            }

            String executionresultpath;

            int indexsheet = 0;

            File dir = new File(resultdirectory);

            if (!dir.exists()) {
                dir.mkdirs();
            }
            executionresultpath = resultdirectory + "/" + datetime("dd_MM_yyyy HHmmss/");
            dir = new File(executionresultpath);

            if (!dir.exists()) {
                dir.mkdirs();
            }
            errorsnappath = executionresultpath + "errorsnap/";
            dir = new File(errorsnappath);
            if (!dir.exists()) {
                dir.mkdirs();
            }

            for (f = 1; f < whitelist.length; f++) {

                FileInputStream ObjExcelApp = new FileInputStream(whitelist[f]);

                if (ObjExcelApp != null) {
                    ObjWorkBook = new HSSFWorkbook(ObjExcelApp);

                    if (ObjWorkBook != null) {
                        HSSFSheet ObjWorkSheet = ObjWorkBook.getSheet(index);
                        resulfilecreation(whitelist[f].getName(), executionresultpath);
                        indexsheet = 0;
                        if (ObjWorkSheet != null) {
                            int totalrows = ObjWorkSheet.getLastRowNum();

                            for (k = 1; k <= totalrows; k++) {
                                {
                                    try {

                                        HSSFRow Row = ObjWorkSheet.getRow(k);

                                        if (Row.getCell(3) != null) {

                                            if (Row.getCell(3).toString().equalsIgnoreCase("Y")) {
                                                indexsheet++;
                                                String sheetname = Row.getCell(2).toString();

                                                for (int sheetpriya = 0; sheetpriya < ObjWorkBook.getNumberOfSheets(); sheetpriya++) {

                                                    String sheeto = ObjWorkBook.getSheetName(sheetpriya);

                                                    if (sheetname.equals(sheeto)) {
                                                        sheet = ObjWorkBook.getSheetAt(sheetpriya);
                                                    }

                                                }

                                                activesheet = sheet.getSheetName().toString();
                                                sheetname = activesheet;
                                                sheetindex = k;

                                                exeresultsheet = exeresultworkbook.createSheet(activesheet);
                                                resultfileheader();
                                                executionstatus = "Pass";
                                                innerarray = new String[sheet.getLastRowNum() + 1][sheet.getRow(0).getLastCellNum()];
                                                // System.out.println("total row"+sheet.getLastRowNum());
                                                for (i = 1; i <= sheet.getLastRowNum(); i++) {

                                                    // {
                                                    if (sheet.getRow(i) != null) {
                                                        Row row = sheet.getRow(i);

                                                        Boolean isEmptyRow = true;
                                                        System.out.println(isEmptyRow);
                                                        for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
                                                            Cell cell = row.getCell(cellNum);
                                                            if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK && StringUtils.isNotBlank(cell.toString())) {
                                                                isEmptyRow = false;

                                                            }

                                                        }
                                                        if (!isEmptyRow) {
                                                            totalrow++;
                                                            System.out.println(isEmptyRow);
                                                            System.out.println(totalrow);
                                                        }
                                                    }
                                                }
                                                for (i = 1; i <= sheet.getLastRowNum(); i++) {

                                                    if (sheet.getRow(i) != null) {
                                                        // innerRowCount = sheet.getRow(i);
                                                        Row row = sheet.getRow(i);
                                                        Boolean isEmptyRow = true;
                                                        HSSFCell innerCell = null;
                                                        int lastColumn = Math.max(9, row.getLastCellNum());
                                                        innercell = new String[totalrow + 1][lastColumn];
                                                        for (int j = 0; j < lastColumn; j++) {

                                                            Cell cell = row.getCell(j);
                                                            if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK && StringUtils.isNotBlank(cell.toString())) {
                                                                isEmptyRow = false;
                                                                System.out.println("in loop");
                                                                cell.setCellType(Cell.CELL_TYPE_STRING);
                                                                String value = cell.toString();
                                                                innerarray[i][j] = value;

                                                            }
                                                        }

                                                        if (!isEmptyRow) {
                                                            if (innerarray[i][6].equalsIgnoreCase("browsertype")) {
                                                                bdriverscheck(innerarray[i][7]);
                                                            }

                                                            if (innerarray[i][6].equalsIgnoreCase("gettext")) {
                                                                rowvalue = i;
                                                            }

                                                            if (innerarray[i][6].equalsIgnoreCase("Search")) {
                                                                skuid = innerarray[i][7];
                                                            }

                                                            if (errorcheck(innerarray[i][6]) == false) {

                                                                try {

                                                                    objCreation(innerarray[i][0], innerarray[i][1], innerarray[i][2], innerarray[i][3], innerarray[i][4], innerarray[i][5],
                                                                            innerarray[i][6], innerarray[i][7], innerarray[i][8]);
                                                                    WritePass(i, innerarray[i][0], innerarray[i][1], innerarray[i][2], errorsnappath, rowvalue);
                                                                    if (innerarray[i][6].equalsIgnoreCase("to")) {
                                                                        try {
                                                                            if (driver.findElement(By.xpath(".//*[@id='ltbx_initMap']/area[3]")).isDisplayed()) {
                                                                                driver.findElement(By.xpath(".//*[@id='ltbx_initMap']/area[3]")).click();
                                                                                System.out.println("closed popup");
                                                                            }
                                                                        } catch (Exception e) {
                                                                            continue;
                                                                        }
                                                                    }

                                                                } catch (Exception e) {

                                                                    e.printStackTrace();
                                                                    if (e.getMessage().contains("For documentation on this error")) {

                                                                        String Str = null;
                                                                        Str = e.getMessage();
                                                                        int startIndex = 0;
                                                                        int index = Str.indexOf("For documentation on this error", startIndex);

                                                                        WriteFail(i, innerarray[i][0], innerarray[i][1], innerarray[i][2], Str.substring(0, index), errorsnappath, false);
                                                                        rowcount = i;
                                                                        executionstatus = "Fail";
                                                                        errorcheck = true;
                                                                        break;

                                                                    } else {

                                                                        WriteFail(i, innerarray[i][0], innerarray[i][1], innerarray[i][2], e.getMessage(), errorsnappath, false);
                                                                        rowcount = i;
                                                                        executionstatus = "Fail";
                                                                        errorcheck = true;
                                                                        break;
                                                                    }

                                                                }

                                                            } else {

                                                                executionstatus = "Not Executed";
                                                                WriteFail(i, innerarray[i][0], innerarray[i][1], innerarray[i][2], errortype, errorsnappath, true);
                                                            }
                                                        }
                                                    }

                                                }

                                            }

                                            exeresultsheet = exeresultworkbook.getSheet(index);
                                            exerow = exeresultsheet.createRow(indexsheet);
                                            execell = exerow.createCell(0);
                                            execell.setCellValue(activesheet);

                                            if (executionstatus.equalsIgnoreCase("Pass")) {
                                                HSSFCellStyle style = exeresultworkbook.createCellStyle();
                                                style.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
                                                style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

                                                execell = exerow.createCell(1);
                                                execell.setCellValue(executionstatus);
                                                execell.setCellStyle(style);
                                            } else if (executionstatus.equalsIgnoreCase("fail")) {
                                                HSSFCellStyle style = exeresultworkbook.createCellStyle();
                                                style.setFillForegroundColor(HSSFColor.ROSE.index);
                                                style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                                                execell = exerow.createCell(1);
                                                execell.setCellValue(executionstatus);
                                                execell.setCellStyle(style);

                                                activesheet = sheet.getSheetName().toString();
                                                exeresultsheet = exeresultworkbook.getSheet(activesheet);

                                                clearcount = rowcount + 2;

                                                int innerrcount = rowcount + 1;
                                                rcount = innerrcount;
                                                System.out.println(innerrcount);
                                                for (rcount = innerrcount; rcount <= totalrow; rcount++) {

                                                    Row row = sheet.getRow(rcount);
                                                    // System.out.println("rcount"+rcount);
                                                    Boolean isEmptyRow = true;

                                                    int lastColumn = Math.max(9, row.getLastCellNum());
                                                    String[][] innercellnotexe = new String[totalrow + 1][lastColumn];
                                                    for (int j = 0; j < lastColumn; j++) { // Column

                                                        Cell cell = row.getCell(j);
                                                        if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK && StringUtils.isNotBlank(cell.toString())) {
                                                            isEmptyRow = false;

                                                        }
                                                        if (!isEmptyRow) {
                                                            System.out.println("in loop" + rcount);
                                                            cell.setCellType(Cell.CELL_TYPE_STRING);
                                                            String value = cell.toString();
                                                            if (value != null) {
                                                                innercellnotexe[rcount][j] = value;
                                                            }
                                                            System.out.println("value" + innercellnotexe[rcount][j]);
                                                        }

                                                    }
                                                    WriteNotExecuted(rcount, innercellnotexe[rcount][0], innercellnotexe[rcount][1], innercellnotexe[rcount][2], true);
                                                    clearcount++;
                                                }
                                                if (totalrow != 0 && clearcount > totalrow) {
                                                    System.out.println("clear cached" + clearcount);
                                                    System.out.println("clear cached" + totalrow);
                                                    ClearBrowserCache();
                                                    clearcount = 0;
                                                }

                                            } else {
                                                HSSFCellStyle style = exeresultworkbook.createCellStyle();
                                                style.setFillForegroundColor(HSSFColor.LEMON_CHIFFON.index);
                                                style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                                                execell = exerow.createCell(1);
                                                execell.setCellValue(executionstatus);
                                                execell.setCellStyle(style);
                                            }
                                            summarycounter++;
                                            totalrow = 0;

                                        }
                                    }

                                    catch (Exception e) {

                                        System.out.println("error" + e);
                                        e.printStackTrace();

                                    }

                                }
                            }

                        }
                    }
                }

                if (errorcheck) {

                    ObjWorkBook = null;
                    ObjExcelApp.close();
                    exeresultworkbook.write(resultexe);
                    resultexe.close();
                    driver.close();
                    driver.quit();
                } else {
                    ObjWorkBook = null;
                    ObjExcelApp.close();
                    exeresultworkbook.write(resultexe);
                    resultexe.close();
                    driver.close();
                    driver.quit();

                }

            }

            // resultexe.flush();

            System.exit(0);
        } catch (Exception e) {
            System.out.println("error in excel read" + e);
            e.printStackTrace();

        }

    }

    // In-case of missing of drivers for IE and Chrome.
    public void bdriverscheck(String btype) {

        if (btype.equalsIgnoreCase("ie")) {
            File bcheck = new File(iedriver);
            if (!bcheck.exists()) {
                JOptionPane.showMessageDialog(null, "Driver File not Exist at: " + iedriver);
                System.exit(0);
            }
        }
        if (btype.equalsIgnoreCase("chrome")) {
            File bcheck = new File(chromedriver);
            if (!bcheck.exists()) {
                JOptionPane.showMessageDialog(null, "Driver File not Exist at: " + chromedriver);
                System.exit(0);
            }
        }

    }

    // Checking error by splitting it by comma from config file.
    public boolean errorcheck(String event) {

        boolean val = false;
        if (event.equalsIgnoreCase("browsertype")) {
            return val;
        } else {
            for (String errorval : error.split(",")) {

                val = driver.getPageSource().contains(errorval);
                if (val) {
                    errortype = errorval;
                    break;
                }
            }
        }

        return val;
    }

    // Mentioning its header for other sheets except Index sheet
    public void resultfileheader() {
        HSSFCellStyle style = exeresultworkbook.createCellStyle();
        style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        exerow = exeresultsheet.createRow(0);
        execell = exerow.createCell(0);
        execell.setCellStyle(style);
        execell.setCellValue("StepId");

        execell = exerow.createCell(1);
        execell.setCellStyle(style);
        execell.setCellValue("TestCase_Type");

        execell = exerow.createCell(2);
        execell.setCellStyle(style);
        execell.setCellValue("Test Case");

        execell = exerow.createCell(3);
        execell.setCellStyle(style);
        execell.setCellValue("ExecutionDateTime");

        execell = exerow.createCell(4);
        execell.setCellStyle(style);
        execell.setCellValue("ExecutionStatus");

        execell = exerow.createCell(5);
        execell.setCellStyle(style);
        execell.setCellValue("Value");

        execell = exerow.createCell(6);
        execell.setCellStyle(style);
        execell.setCellValue("ErrorLog");

        execell = exerow.createCell(7);
        execell.setCellStyle(style);
        execell.setCellValue("ExecutionSnap");

        execell = exerow.createCell(8);
        execell.setCellStyle(style);
        execell.setCellValue("Text Found");

    }

    // Creating result file creation and mentioning its header for Index sheet
    public void resulfilecreation(String filename, String filepath) {

        try {
            Date currentdate = new Date();
            SimpleDateFormat dateformat = new SimpleDateFormat("ddMMyyyy_HHmmss");
            exeresultworkbook = new HSSFWorkbook();

            resultexe = new FileOutputStream(filepath + dateformat.format(currentdate) + filename);
            exeresultsheet = exeresultworkbook.createSheet(index);
            HSSFCellStyle style = exeresultworkbook.createCellStyle();
            style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
            style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            exerow = exeresultsheet.createRow(0);
            execell = exerow.createCell(0);
            execell.setCellStyle(style);
            execell.setCellValue("Scenario");

            execell = exerow.createCell(1);
            execell.setCellStyle(style);
            execell.setCellValue("Status");
        } catch (Exception e) {
            System.out.println("error in result file creation block" + e);
            e.printStackTrace();
        }

    }

    // Various enum for the Object type available
    public enum objecType {
        Browser, TextBox, Button, Label, Link, Custom, RadioButton, MouseHover, CheckBox, DropDown, CustomValue;
    }

    // Various enum for the Object event available
    public enum objecEvent {
        Clear, Click, IsVisible, IsSelected, IsTextPresent, Type, Launch, Close, BrowserType, PromoPopup, Custom, Submit, GetText, TO, FORWARD, BACKWARD, CurrentURL, MouseHover, VerifyStatus, Login, LoginValidation, Logout, Search, LoginCheckout, Wait, iFrame, SelectByValue, SelectByText, SelectByIndex, Popup, PopupClose, Alert, FileUpload, DefaultWindow, iFrameTextBoxType, switchDefault, iFrameButtonClick;

    }

    public enum objecValue {
        Emailid;
    }

    @SuppressWarnings("incomplete-switch")
    // Going to the case which is called by the test sheet.
    public void objCreation(String pStrStepID, String TestCase_Type, String pStrTestCaseName, String pStrObjectType, String pStrObjectName, String pStrObjectDescription, String pStrEvent,
            String pStrValue, String pStrComment) throws Exception

    {
        if (pStrValue != null) {
            if (pStrValue.equals("Emailid")) {
                pStrValue = CustomMethods.emailId();
            }
        }

        objecType objType = null;
        if (pStrObjectType != null) {
            objType = objecType.valueOf(pStrObjectType);
        }

        objecEvent objevent = null;
        if (pStrEvent != null) {
            objevent = objecEvent.valueOf(pStrEvent);
        }

        switch (objType)

        {

            case Custom: {

                switch (objevent) {
                    case Login: {

                        CustomMethods.login(driver, pStrValue);
                        break;
                    }
                    case LoginValidation: {

                        CustomMethods.loginValidation(driver, pStrValue);
                        break;
                    }
                    case Logout: {
                        pStrValue = CustomMethods.logout(driver);
                        break;
                    }
                    case Search: {
                        CustomMethods.search(driver);
                        break;
                    }
                    case LoginCheckout: {
                        CustomMethods.loginCheckout(driver, pStrValue);
                        break;
                    }
                    case Wait: {
                        waitUntilElementIsPresent(pStrObjectName, pStrObjectDescription);
                        break;
                    }
                    case iFrame: {
                        switchToiFrame(pStrValue);
                        break;
                    }
                    case iFrameTextBoxType: {

                        iFrameTextBoxType(pStrObjectName, pStrObjectDescription, pStrValue);
                        break;

                    }
                    case iFrameButtonClick: {

                        iFrameButtonClick(pStrObjectName, pStrObjectDescription);
                        break;

                    }
                    case FileUpload: {
                        CustomMethods.fileUpload();
                        break;
                    }
                    case switchDefault: {
                        switchToDefault();
                        break;
                    }
                }
                break;
            }

            case MouseHover: {

                switch (objevent) {
                    case MouseHover: {

                        mouseHover(pStrObjectName, pStrObjectDescription);
                        break;
                    }

                }
                break;
            }
            case Browser: {

                switch (objevent) {
                    case BrowserType: {
                        browserType(pStrValue);
                        break;
                    }
                    case Launch: {
                        openbrowser(pStrValue);
                        break;
                    }
                    case Close: {
                        closebrowser();
                        break;
                    }
                    case PromoPopup: {
                        promotionalPopupCheck(pStrObjectDescription);
                        break;
                    }
                    case TO: {
                        To(pStrObjectDescription);
                        break;
                    }
                    case FORWARD: {
                        Forward();
                        break;
                    }
                    case BACKWARD: {
                        Backward();
                        break;
                    }
                    case CurrentURL: {
                        GetCurrentUrl(driver);
                        break;
                    }

                    case VerifyStatus: {
                        // verifyURLStatus();
                        break;
                    }
                    case Popup: {
                        popUpHandle(pStrObjectDescription);
                        break;
                    }
                    case PopupClose: {
                        popUpHandleClose();
                        break;
                    }
                    case Alert: {
                        alertHandle();
                        break;
                    }
                    case DefaultWindow: {
                        defaultWindow();
                        break;
                    }
                }
                break;
            }

            case TextBox: {

                switch (objevent) {
                    case Clear: {
                        TextBoxClear(pStrObjectName, pStrObjectDescription);
                        break;
                    }
                    case Type: {
                        TextBoxType(pStrObjectName, pStrObjectDescription, pStrValue);
                        break;
                    }
                    case Click: {
                        Click(pStrObjectName, pStrObjectDescription);
                        break;
                    }

                    case Submit: {
                        Submit(pStrObjectName, pStrObjectDescription);
                        break;
                    }

                    case IsVisible: {
                        IsVisible(pStrObjectName, pStrObjectDescription);
                        break;
                    }
                    case GetText: {
                        getText(pStrObjectName, pStrObjectDescription);
                        break;
                    }
                }
                break;
            }
            case Button: {
                switch (objevent) {

                    case Click: {
                        Click(pStrObjectName, pStrObjectDescription);

                        break;
                    }
                    case IsVisible: {
                        IsVisible(pStrObjectName, pStrObjectDescription);
                        break;
                    }

                }
                break;
            }

            case Link: {
                switch (objevent) {

                    case Click: {

                        Click(pStrObjectName, pStrObjectDescription);
                        break;
                    }

                    case IsVisible: {
                        IsVisible(pStrObjectName, pStrObjectDescription);
                        break;
                    }
                }
                break;
            }

            case Label: {
                switch (objevent) {

                    case IsTextPresent: {
                        IsTextPresent(pStrObjectName, pStrObjectDescription, pStrValue);
                        break;
                    }
                    case IsVisible: {
                        IsVisible(pStrObjectName, pStrObjectDescription);
                        break;
                    }
                }
                break;
            }

            case RadioButton: {
                switch (objevent) {

                    case Click: {
                        Click(pStrObjectName, pStrObjectDescription);
                        break;
                    }
                    case IsSelected: {
                        IsSelected(pStrObjectName, pStrObjectDescription);
                        break;
                    }
                }
            }
            case CheckBox: {
                switch (objevent) {

                    case Click: {
                        Click(pStrObjectName, pStrObjectDescription);
                        break;
                    }
                    case IsSelected: {
                        IsSelected(pStrObjectName, pStrObjectDescription);
                        break;
                    }
                }
                break;
            }
            case DropDown: {
                switch (objevent) {

                    case SelectByValue: {
                        SelectByValue(pStrObjectName, pStrObjectDescription, pStrValue);
                        break;
                    }
                    case SelectByText: {
                        SelectByVisibleText(pStrObjectName, pStrObjectDescription, pStrValue);
                        break;
                    }
                    case SelectByIndex: {
                        SelectByIndex(pStrObjectName, pStrObjectDescription, pStrValue);
                        break;
                    }
                }
                break;
            }
        }

    }

    public void DateTIme() {
        try {
            HSSFCell date_time;
            date_time = sheet.getRow(i).createCell(8);
            date = new Date();
            date_time.setCellValue(date.toString());

            FileOutputStream out = new FileOutputStream(Driver.InputExcelPath[f]);
            ObjWorkBook.write(out);
            out.flush();
            out.close();
        } catch (Exception e) {
            System.out.println("error in date time block" + e);
            e.printStackTrace();
        }
    }

    // Reading config file parameteres.
    public static void config(File configfile) {

        try {
            String[][] innerarray = null;
            HSSFSheet sheetc = null;
            HSSFCell innerCell = null;

            FileInputStream ObjExcelconfig = new FileInputStream(configfile);
            if (ObjExcelconfig != null) {
                HSSFWorkbook ObjWorkBookconfig = new HSSFWorkbook(ObjExcelconfig);
                if (ObjWorkBookconfig != null) {
                    HSSFSheet configsheet = ObjWorkBookconfig.getSheet("Config");

                    ArrayList<String> obj = new ArrayList<String>();
                    for (int i = 1; i <= configsheet.getLastRowNum(); i++) {
                        HSSFRow Row = configsheet.getRow(i);

                        innerCell = Row.getCell(1);
                        innerCell.setCellType(Cell.CELL_TYPE_STRING);

                        obj.add(innerCell.toString());
                    }

                    index = obj.get(0);
                    resultdirectory = obj.get(1);
                    errorsnapshotforpass = obj.get(2);
                    errorsnapshotforfail = obj.get(3);
                    chromedriver = obj.get(4);
                    iedriver = obj.get(5);
                    error = obj.get(6);

                }
            }
        } catch (Exception e) {
            System.out.println("error in config" + e);
            e.printStackTrace();

        }
    }
}
