package com.otc.selenium;

import java.util.Random;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class CustomMethods extends GenericMethods {

    public CustomMethods(Driver d) {
        super(d);
        // TODO Auto-generated constructor stub
    }

    public static void login(WebDriver driver, String credential) throws Exception {
        // Pattern pattern ;
        String emailid = null;
        String password = null;
        // pattern = Pattern.compile("\\W");
        // Matcher m = pattern.matcher(credential);
        // System.out.println("Split using Pattern.split(): " + m.group());
        // String[] words = pattern.split(credential);
        // for (String s : words) {
        // System.out.println("Split using Pattern.split(): " + s);
        // }
        // "(.*)(\\d+)(.*)"
        // String REGEX = "^(.*)(\\W+)(.*)";
        // System.out.println(REGEX);
        // Pattern p = Pattern.compile(REGEX);
        // get a matcher object
        // Matcher m = p.matcher(credential);
        // m.group();

        String[] logincredential = credential.split(",");
        emailid = logincredential[0];
        password = logincredential[1];
        System.out.println(emailid);
        System.out.println(password);
        Actions action = new Actions(driver);
        WebElement element = driver.findElement(By.xpath("//header[@class='body-header']/div[2]/div/nav/ul/li[1]/button[2]"));
        action.moveToElement(element).build().perform();
        driver.findElement(By.xpath("//header[@class='body-header']/div[2]/div/nav/ul/li[1]/ul/li[1]/a")).click();
        waitUntilElementIsPresent("xpath", "//div[@class='main-wrapper']/article/header/h1");
        WebElement email = driver.findElement(By.id("userId"));
        email.sendKeys(emailid);
        WebElement pass = driver.findElement(By.id("userPassword"));
        pass.sendKeys(password);
        driver.findElement(By.xpath("//form[@id='frmOrderLogin1']/button")).submit();
    }

    public static void loginValidation(WebDriver driver, String credential) throws Exception {

        System.out.println("return");
        String emailid = null;
        String password = null;
        String validationtype = null;
        Boolean flagloginheader;
        WebDriverWait wait = new WebDriverWait(driver, 15);
        String[] logincredential = credential.split(",");
        String label1;
        if (logincredential[0].contains("@") && logincredential[0].contains(".")) {
            emailid = logincredential[0];
            validationtype = logincredential[1];
            System.out.println(logincredential[1]);
        } else {
            password = logincredential[0];
            validationtype = logincredential[1];
            System.out.println(logincredential[0]);
        }

        if (driver.findElement(By.xpath("//header[@class='body-header']/div[2]/div/nav/ul/li[1]/button[2]")).isEnabled()) {
            Actions action = new Actions(driver);

            WebElement element = driver.findElement(By.xpath("//header[@class='body-header']/div[2]/div/nav/ul/li[1]/button[2]"));
            action.moveToElement(element).build().perform();
            driver.findElement(By.xpath("//header[@class='body-header']/div[2]/div/nav/ul/li[1]/ul/li[1]/a")).click();
            waitUntilElementIsPresent("xpath", "//div[@class='main-wrapper']/article/header/h1");
            flagloginheader = driver.findElement(By.xpath("//div[@class='main-wrapper']/article/header/h1")).isDisplayed();
            if (flagloginheader == true) {
                label1 = driver.findElement(By.xpath("//div[@class='main-wrapper']/article/header/h1")).getText();
                if (label1.equalsIgnoreCase("Log In or Create an Account")) {

                    WebElement email = driver.findElement(By.id("userId"));
                    if (email.isDisplayed()) {
                        if (emailid != null) {
                            email.sendKeys(emailid);
                        }
                        System.out.println(emailid);
                    }
                    WebElement pass = driver.findElement(By.id("userPassword"));
                    if (pass.isDisplayed()) {
                        if (password != null) {
                            pass.sendKeys(password);
                            System.out.println(password);
                        }
                    }

                    if (driver.findElement(By.xpath("//div[@class='module-wrapper']/form/button")).isDisplayed()) {
                        driver.findElement(By.xpath("//div[@class='module-wrapper']/form/button")).submit();
                    }

                    label1 = driver.findElement(By.xpath("//div[@class='main-wrapper']/article/header/h1")).getText();
                    if (label1.equalsIgnoreCase("Log In or Create an Account") && validationtype.equals("email")) {
                        validationempassword(driver);
                    } else if (label1.equalsIgnoreCase("Log In or Create an Account") && validationtype.equals("password")) {
                        validationemail(driver);
                    } else if (label1.equalsIgnoreCase("Log In or Create an Account") && validationtype.equals("both")) {
                        validationboth(driver);
                    }

                    else if (label1.equalsIgnoreCase("Forgot Your Password?")) {
                        System.out.println("forgot password");
                        loginthreeattemptvalidation(driver);
                    }

                    else {

                        throw new Exception("Expected Value not match with Actual Value" + driver.findElement(By.xpath("html/body/div[2]/main/div/article/header/h1")).getText());
                    }

                }

            }
        }

    }

    public static void loginCheckout(WebDriver driver, String credential) throws Exception {

        String label2;
        String emailid = null;
        String password = null;
        Boolean flagloginreturn;
        String[] logincredential = credential.split(",");

        if (logincredential.length == 2) {

            emailid = logincredential[0];
            password = logincredential[1];
        } else if (logincredential.length == 1) {
            if (logincredential[0].contains("@") || logincredential[0].contains(".")) {
                emailid = logincredential[0];

            } else {
                password = logincredential[0];

            }

        }

        label2 = driver.findElement(By.xpath("html/body/div[2]/div/div[1]/div/div/section/div/div[2]/h2")).getText();
        System.out.println(label2);
        if (label2.equalsIgnoreCase("Returning Customers")) {

            WebElement email = driver.findElement(By.id("userId"));
            if (email.isDisplayed()) {
                System.out.println("email");
                if (emailid != null) {
                    System.out.println("not null");
                    email.sendKeys(emailid);
                }
            }
            WebElement pass = driver.findElement(By.id("userPassword"));
            if (pass.isDisplayed()) {
                System.out.println("password");
                if (password != null) {
                    System.out.println("not null pass");
                    pass.sendKeys(password);

                }

                driver.findElement(By.xpath(".//*[@id='frmOrderLogin1']/input[2]")).submit();
            }

            else {

                throw new Exception("Expected Value not match with Actual Value");
            }

        }

    }

    public static void validationemail(WebDriver driver) throws Exception {

        waitUntilElementIsPresent("id", "userId-error");
        System.out.println(driver.findElement(By.id("userId-error")).getText());
        throw new Exception("Please provide a valid email address");
    }

    public static void validationempassword(WebDriver driver) throws Exception {
        waitUntilElementIsPresent("id", "userPassword-error");
        System.out.println(driver.findElement(By.id("userPassword-error")).getText());
        throw new Exception("This field is required.");
    }

    public static void validationboth(WebDriver driver) throws Exception {
        waitUntilElementIsPresent("id", "userId-error");
        waitUntilElementIsPresent("id", "userPassword-error");
        throw new Exception("Please provide a valid email address and password");
    }

    public static void loginthreeattemptvalidation(WebDriver driver) throws Exception {
        String value = "Your account has been locked temporarily after 3 attempts to login";
        System.out.println(driver.findElement(By.xpath(".//*[@id='frmPasswordAssist']/div")).getText());
        System.out.println(driver.findElement(By.xpath(".//*[@id='frmPasswordAssist']/div")).getText().contains(value));
        if (driver.findElement(By.xpath(".//*[@id='frmPasswordAssist']/div")).isDisplayed() && driver.findElement(By.xpath(".//*[@id='frmPasswordAssist']/div")).getText().contains(value)) {
            System.out.println(driver.findElement(By.xpath(".//*[@id='frmPasswordAssist']/div")).getText());
            throw new Exception("Your account has been locked temporarily after 3 attempts to login.Please wait 30 minutes to try again or reset your password by entering your email address below");

        } else {
            throw new Exception("Expected value Your account has been locked temporarily after 3 attempts to login. not matched with"
                    + driver.findElement(By.xpath(".//*[@id='frmPasswordAssist']/div")).getText());
        }
    }

    public static String logout(WebDriver driver) {

        Actions action = new Actions(driver);
        WebElement element = driver.findElement(By.xpath("//header[@class='body-header']/div[2]/div/nav/ul/li[1]/button[2]"));
        action.moveToElement(element).build().perform();
        driver.findElement(By.xpath("//header[@class='body-header']/div[2]/div/nav/ul/li[1]/ul/li[9]/a")).click();
        String text = driver.findElement(By.id("global-content-wrap")).getText();
        return text;

    }

    public static void search(WebDriver driver) {

        driver.findElement(By.name("Ntt")).clear();
        WebElement skuid = driver.findElement(By.name("Ntt"));
        skuid.sendKeys(GenericMethods.skuid);
        driver.findElement(By.name("Ntt")).submit();

    }

    public static void fileUpload() {

    }

    public static String emailId() {
        Random rn = new Random();
        int value = rn.nextInt();
        String append = Integer.toString(value);
        System.out.println("randomvalue" + append);
        String emailname = "testemail";
        String emailid = emailname + append + "@" + "gmail.com";
        return emailid;

    }

}
