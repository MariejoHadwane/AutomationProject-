import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.interactions.Actions;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.concurrent.TimeUnit;

public class FlowTest {

    private WebDriver driver;
    private Actions actions;

    public void setDriver(String driverPath) {
        System.setProperty("webdriver.chrome.driver", driverPath);
        driver = new ChromeDriver();
        actions = new Actions(this.driver);
    }

    public void closeDriver(){

        driver.close();
    }

    private WebElement getWaitForClickableElement(By byLocator){
        WebDriverWait wait = new WebDriverWait(driver, 1000);
        return wait.until(ExpectedConditions.elementToBeClickable(byLocator));
    }

    private void sendKeys(By byLocator, CharSequence keysToSend)
    {
        driver.findElement(byLocator).sendKeys(keysToSend);
    }

    private void hoverOverElement(By byLocator){
        WebElement element = this.getWaitForClickableElement(byLocator);
        actions.moveToElement(element).perform();
    }

    private void clickOnElement(By byLocator){
        WebElement element = this.getWaitForClickableElement(byLocator);
        actions.click(element).perform();
    }

    private void doubleClickOnElement(By byLocator){
        WebElement element = this.getWaitForClickableElement(byLocator);
        actions.doubleClick(element).perform();
    }
    //scrolldown
    private void executeScript(By byLocator, String script){
        JavascriptExecutor javascriptExecutor = (JavascriptExecutor) driver;
        WebElement element = driver.findElement(byLocator);
        javascriptExecutor.executeScript(script, element);
    }

    public void goToUrl(String url){
        this.driver.get(url);
    }

    public void login(String username, String password){
        this.clickOnElement(By.id("acceptCookie"));
        this.sendKeys(By.id("username"), username);
        this.sendKeys(By.id("mot_passe"), password);
        this.clickOnElement(By.id("logBtn"));
    }

    public void testMiseAjourMassive(){
        //Step 1: Login
        try {
            this.goToUrl("https://test.flow.com");
            driver.manage().window().maximize();
            this.login("Username", "Password");
        } catch (Exception e) {
            System.out.println("ERROR: Login. Message: " + e.getMessage());
        }

        //Step 2: Cliquer sur 'Donnees'
        try {
            this.hoverOverElement(By.id("menu_donnees"));
        } catch (Exception e) {
            System.out.println("ERROR: Cliquer sur 'Donnees'. Message: " + e.getMessage());
        }

        //Step 3: Naviguer vers la page Mise a jour massive
        try {
            this.doubleClickOnElement(By.id("integration-autonome"));
        } catch (Exception e) {
            System.out.println("ERROR: Naviguer vers la page Mise a jour massive. Message: " + e.getMessage());
        }

        //Step 4: Cliquer sur 'Creer un fichier'
        try {
            this.clickOnElement(By.cssSelector("body > app-root > div > app-screens > app-header > div > div > div > app-integration-autonome > div > div > div.breadCrumbsRow.row > div.col.buttonsDiv > span > button:nth-child(2)"));
        } catch (Exception e) {
            System.out.println("ERROR: Cliquer sur 'Creer un fichier'. Message: " + e.getMessage());
        }

        //Step 5: Choisir la valeur 'Identifiants' dans le champs 'Module a extraire'
        try {
            this.clickOnElement(By.xpath("//*[@id=\'configurationCard\']/div[2]/div[2]/div/div[3]/app-field-multi-select/div/button"));
            this.clickOnElement(By.cssSelector("#multiFieldModules > li:nth-child(1)"));
        } catch (Exception e) {
            System.out.println("ERROR: Choisir la valeur 'Identifiants' dans le champs 'Module a extraire'. Message: " + e.getMessage());
        }

        //Step 6: Choisir la valeur 'Libelle court' dans le champs 'Champs a extraire'
        try {
            this.clickOnElement(By.cssSelector("div.form-group:nth-child(5) > app-field-multi-select:nth-child(2) > div:nth-child(1) > button:nth-child(1)"));
            By locatorLibelleCourt = By.cssSelector("li.check-list-item:nth-child(18) > input:nth-child(1)");
            String scrollDown = "arguments[0].scrollIntoView(true);";
            this.executeScript(locatorLibelleCourt, scrollDown);
            this.clickOnElement(locatorLibelleCourt);
        } catch (Exception e) {
            System.out.println("ERROR: Choisir la valeur 'Libelle court' dans le champs 'Champs a extraire'. Message: " + e.getMessage());
        }

        //Step 7: Choisir les langues 'Francais' et 'Anglais'
        try {
            this.clickOnElement(By.cssSelector("div.form-group:nth-child(6) > app-field-multi-select:nth-child(2) > div:nth-child(1) > button:nth-child(1)"));
            this.clickOnElement(By.cssSelector("li.check-list-item:nth-child(1)"));
        } catch (Exception e) {
            System.out.println("ERROR: Choisir les langues 'Francais' et 'Anglais'. Message: " + e.getMessage());
        }

        //Step 8: Cliquer 'Continuer'
        try {
            this.clickOnElement(By.cssSelector("button.btn-primary:nth-child(2)"));
        } catch (Exception e) {
            System.out.println("ERROR: Cliquer 'Continuer' (1). Message: " + e.getMessage());
        }

        //Step 9: Saisir dans 'Libelle Court' le mot 'Test'
        try {
            this.sendKeys(By.cssSelector("input.form-control:nth-child(3)"), "Test");
        } catch (Exception e) {
            System.out.println("ERROR: Saisir dans 'Libelle Court' le mot 'Test'. Message: " + e.getMessage());
        }

        //Step 10: Cliquer 'Continuer'
        try {
            this.clickOnElement(By.cssSelector("button.btn-primary:nth-child(2)"));
        } catch (Exception e) {
            System.out.println("ERROR: Cliquer 'Continuer' (2). Message: " + e.getMessage());
        }

        //Step 11: Selectionner les 3 premiers produits de la grille
        try {
            String cssSelector;
            for (int i = 1; i < 4; ++i) {
                cssSelector = "tr.ng-star-inserted:nth-child(" + i + ") > td:nth-child(1) > input:nth-child(1)";
                this.clickOnElement(By.cssSelector(cssSelector));
            }
        } catch (Exception e) {
            System.out.println("ERROR: Selectionner les 3 premiers produits de la grille. Message: " + e.getMessage());
        }

        //Step 12: Generer le fichier
        try {
            this.clickOnElement(By.cssSelector("button.float-right:nth-child(2)"));
        } catch (Exception e) {
            System.out.println("ERROR: Generer le fichier. Message: " + e.getMessage());
        }
    }

    public void testExcel() {
        try {
            String dateToday = new SimpleDateFormat("yyyy-M-dd").format(Calendar.getInstance().getTime());
            FileInputStream fis = new FileInputStream("C:\\Users\\Lenovo\\Downloads\\ExcelFile.xlxs" + dateToday + ".xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheet("Produit");
            XSSFRow row = sheet.getRow(0);

            int numberOfColumns = row.getLastCellNum();
            XSSFCell cellChamp, cellLangue;
            String cellChampValue, cellLangueValue;
            for (int i = 0; i < numberOfColumns; ++i) {
                cellChamp = sheet.getRow(1).getCell(i);
                cellChampValue = cellChamp.getStringCellValue();

                cellLangue = sheet.getRow(5).getCell(i);
                cellLangueValue = cellLangue.getStringCellValue();

                if (cellChampValue.equals("Libellé court") && cellLangueValue.equals("Français")) {
                    XSSFCell cellLibelleCourtFrancais;
                    String cellLibelleCourtFrancaisValue;
                    XSSFCellStyle cellLibelleCourtFrancaisStyle;
                    Color cellLibelleCourtFrancaisColor;
                    for (int j = 6; j < 9; j++) {
                        cellLibelleCourtFrancais = sheet.getRow(j).getCell(i);
                        cellLibelleCourtFrancaisValue = cellLibelleCourtFrancais.getStringCellValue();
                        Assert.assertEquals("ERROR: Valeur Libelle Court", "Test", cellLibelleCourtFrancaisValue);

                        cellLibelleCourtFrancaisStyle = cellLibelleCourtFrancais.getCellStyle();
                        cellLibelleCourtFrancaisColor = cellLibelleCourtFrancaisStyle.getFillForegroundXSSFColor();
                        Assert.assertEquals("ERROR: Couleur Libelle Court", "FFFFFF00", ((XSSFColor) cellLibelleCourtFrancaisColor).getARGBHex());
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("ERROR Excel. Message: " + e.getMessage());
        }
    }

    public static void main(String[] args) {
        String driverPath = "C:\\Users\\Lenovo\\Desktop\\chromedriver.exe";
        FlowTest flowTest = new FlowTest();
        flowTest.setDriver(driverPath);
        flowTest.testMiseAjourMassive();
        try {
            TimeUnit.SECONDS.sleep(30);
        } catch (Exception e) {

        }
        flowTest.closeDriver();
        flowTest.testExcel();
    }
}