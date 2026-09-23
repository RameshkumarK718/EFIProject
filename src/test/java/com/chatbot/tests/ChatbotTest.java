package com.chatbot.tests;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

public class ChatbotTest {

    // EXTENDED TEST DATA MODEL (Supports columns 0 to 12)
    public static class TestRowData {
        public String sheetName;
        public int rowIndex;
        public String col0; // Test Case ID / Primary Identifier
        public String col1; // Category
        public String col2; // Subcategory
        public String col3; // Question / Input Prompt
        public String col4; // Expected Answer
        public String col5; // Extra Field / Context
        public String col6; // Extra Field / Priority
        public String col7; // Extra Field / Tags
        public String col8; // Extra Field / Notes
        public String col9; // Runnable Flag (YES/NO)
        public String col10; // Chatbot Actual Answer Output
        public String col11; // Failure Reason / Audit Log
        public String col12; // Final Test Status (PASS/FAIL)

        public TestRowData(String sheetName, int rowIndex, String col0, String col1, String col2,
                           String col3, String col4, String col5, String col6, String col7,
                           String col8, String col9, String col10, String col11, String col12) {
            this.sheetName = sheetName;
            this.rowIndex = rowIndex;
            this.col0 = col0;
            this.col1 = col1;
            this.col2 = col2;
            this.col3 = col3;
            this.col4 = col4;
            this.col5 = col5;
            this.col6 = col6;
            this.col7 = col7;
            this.col8 = col8;
            this.col9 = col9;
            this.col10 = col10;
            this.col11 = col11;
            this.col12 = col12;
        }
    }

    // ==================== UNIVERSAL CONFIGURATIONS ====================
    private static final Duration WAIT_TIMEOUT = Duration.ofSeconds(45);
    private static final Duration RESPONSE_TIMEOUT = Duration.ofSeconds(45);

    // Easily override via JVM arguments: -Dapp.url=... -Dframework.excel.url=... -Dcredentials.excel.url=...
    private static final String APP_URL = System.getProperty("app.url", "https://d3rl0fkw0q6ssb.cloudfront.net/");
    private static final String CREDENTIALS_EXCEL_URL = System.getProperty("credentials.excel.url", 
            "https://raw.githubusercontent.com/RameshkumarK718/Chatbot-Automation-Framework/main/credentials(1).xlsx");
    private static final String FRAMEWORK_EXCEL_URL = System.getProperty("framework.excel.url", 
            "https://raw.githubusercontent.com/RameshkumarK718/Chatbot-Automation-Framework/main/Frameworks(EFI).xlsx");

    private static final String OUTPUT_EXCEL_FILE = "Frameworks_Output.xlsx";

    // ==================== COLUMN MAPPINGS (0 to 12 Index Map) ====================
    private static final int COL_ID = 0;
    private static final int COL_CATEGORY = 1;
    private static final int COL_SUBCATEGORY = 2;
    private static final int COL_QUESTION = 3;
    private static final int COL_EXPECTED_ANSWER = 4;
    private static final int COL_EXTRA_5 = 5;
    private static final int COL_EXTRA_6 = 6;
    private static final int COL_EXTRA_7 = 7;
    private static final int COL_EXTRA_8 = 8;
    private static final int COL_RUNNABLE = 9;
    private static final int COL_CHATBOT_ANSWER = 10;
    private static final int COL_REASON = 11;
    private static final int COL_STATUS = 12;

    private WebDriver driver;
    private WebDriverWait wait;

    // GET EXCEL INPUT STREAM
    private InputStream openUrlStream(String fileUrl) throws Exception {
        URL url = new URL(fileUrl);
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("GET");
        connection.setConnectTimeout(30000);
        connection.setReadTimeout(60000);
        connection.setRequestProperty("User-Agent", "Mozilla/5.0");
        connection.setRequestProperty("Accept", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        int responseCode = connection.getResponseCode();
        if (responseCode != HttpURLConnection.HTTP_OK) {
            throw new RuntimeException("Unable to download file. HTTP response code: " + responseCode + " | URL: " + fileUrl);
        }
        return connection.getInputStream();
    }

    // READ CREDENTIALS FROM EXCEL
    private String[] getCredentialsFromExcel() {
        String username = "";
        String password = "";
        try (InputStream is = openUrlStream(CREDENTIALS_EXCEL_URL);
             Workbook workbook = new XSSFWorkbook(is)) {
            if (workbook.getNumberOfSheets() == 0) {
                throw new RuntimeException("Credentials Excel contains no sheets.");
            }
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(3); // Row 4 (0-indexed)
            if (row == null) {
                throw new RuntimeException("Credentials row 4 was not found in Excel.");
            }
            username = getCellStringValue(row.getCell(0));
            password = getCellStringValue(row.getCell(1));
            System.out.println("--> Credentials loaded successfully.");
        } catch (Exception e) {
            throw new RuntimeException("Error reading credentials Excel: " + e.getMessage(), e);
        }
        return new String[]{username, password};
    }

    private String getCellStringValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        try {
            return formatter.formatCellValue(cell).trim();
        } catch (Exception e) {
            return cell.toString().trim();
        }
    }

    // UNIVERSAL 0-12 EXCEL DATA FETCHER
    private List<TestRowData> fetchExcelDataFromGitHub(String fileUrl) {
        List<TestRowData> dataList = new ArrayList<>();

        try (InputStream is = openUrlStream(fileUrl);
             Workbook workbook = new XSSFWorkbook(is)) {
            if (workbook.getNumberOfSheets() == 0) {
                throw new RuntimeException("Framework Excel contains no sheets.");
            }
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();
                System.out.println("--> Reading sheet: " + sheetName);

                for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) {
                        continue;
                    }

                    String testCaseId = getCellStringValue(row.getCell(COL_ID));
                    String category = getCellStringValue(row.getCell(COL_CATEGORY));
                    String subcategory = getCellStringValue(row.getCell(COL_SUBCATEGORY));
                    String question = getCellStringValue(row.getCell(COL_QUESTION));
                    String expectedAnswer = getCellStringValue(row.getCell(COL_EXPECTED_ANSWER));
                    String extra5 = getCellStringValue(row.getCell(COL_EXTRA_5));
                    String extra6 = getCellStringValue(row.getCell(COL_EXTRA_6));
                    String extra7 = getCellStringValue(row.getCell(COL_EXTRA_7));
                    String extra8 = getCellStringValue(row.getCell(COL_EXTRA_8));
                    String runnable = getCellStringValue(row.getCell(COL_RUNNABLE));

                    // Skip row if marked as NO in the runnable column
                    if ("NO".equalsIgnoreCase(runnable)) {
                        continue;
                    }
                    if (question.isBlank()) {
                        continue;
                    }
                    if (testCaseId.isBlank()) {
                        testCaseId = String.format("TC-%03d", r);
                    }

                    dataList.add(new TestRowData(
                            sheetName, r, testCaseId, category, subcategory,
                            question, expectedAnswer, extra5, extra6, extra7,
                            extra8, runnable, "", "", ""
                    ));
                }
            }
            System.out.println("--> Successfully parsed " + dataList.size() + " test cases (Columns 0-12 supported).");
        } catch (Exception e) {
            throw new RuntimeException("Error fetching Frameworks Excel: " + e.getMessage(), e);
        }
        return dataList;
    }

    @BeforeMethod
    public void setUp() {
        initializeDriverAndLogin();
    }

    @AfterMethod
    public void tearDown() {
        if (driver != null) {
            try {
                driver.quit();
                System.out.println("--> WebDriver closed successfully.");
            } catch (Exception e) {
                System.err.println("--> Error closing WebDriver: " + e.getMessage());
            }
        }
    }

    private void initializeDriverAndLogin() {
        String[] credentials = getCredentialsFromExcel();
        String memberId = credentials[0];
        String password = credentials[1];
        Assert.assertFalse(memberId.isBlank(), "Member ID is missing from Cloud Excel.");
        Assert.assertFalse(password.isBlank(), "Password is missing from Cloud Excel.");

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--headless=new");
        options.addArguments("--no-sandbox");
        options.addArguments("--disable-dev-shm-usage");
        options.addArguments("--disable-gpu");
        options.addArguments("--window-size=1920,1080");
        options.addArguments("--remote-allow-origins=*");

        driver = new ChromeDriver(options);
        driver.manage().timeouts().implicitlyWait(Duration.ZERO);
        driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(60));
        wait = new WebDriverWait(driver, WAIT_TIMEOUT);

        try {
            System.out.println("--> Opening application URL: " + APP_URL);
            driver.get(APP_URL);

            WebElement memberInput = wait.until(ExpectedConditions.elementToBeClickable(By.id("vaa-email")));
            memberInput.clear();
            memberInput.sendKeys(memberId);

            WebElement passwordInput = wait.until(ExpectedConditions.elementToBeClickable(By.id("vaa-pw")));
            passwordInput.clear();
            passwordInput.sendKeys(password);

            WebElement loginButton = wait.until(ExpectedConditions.elementToBeClickable(By.id("vaa-submit")));
            clickElement(loginButton);

            wait.until(ExpectedConditions.or(
                    ExpectedConditions.visibilityOfElementLocated(By.id("vaa-portal")),
                    ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(normalize-space(),'Conversational AI')]"))
            ));

            WebElement conversationalAILink = wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("//span[contains(normalize-space(),'Conversational AI')]/ancestor::a[1] | //a[contains(normalize-space(),'Conversational AI')]")));
            clickElement(conversationalAILink);

            waitForChatInput();
            System.out.println("--> Chatbot input is ready.");
        } catch (Exception e) {
            throw new AssertionError("Failed to initialize chatbot test setup: " + e.getMessage(), e);
        }
    }

    private By getChatInputLocator() {
        return By.xpath("//input[@placeholder='Ask a question...'] | //textarea[@placeholder='Ask a question...'] | //input[contains(@placeholder,'Ask')] | //textarea[contains(@placeholder,'Ask')] | //div[@contenteditable='true']");
    }

    private WebElement waitForChatInput() {
        return wait.until(ExpectedConditions.elementToBeClickable(getChatInputLocator()));
    }

    private void clickElement(WebElement element) {
        try {
            element.click();
        } catch (Exception e) {
            JavascriptExecutor js = (JavascriptExecutor) driver;
            js.executeScript("arguments[0].scrollIntoView({block:'center'});", element);
            js.executeScript("arguments[0].click();", element);
        }
    }

    private List<WebElement> getChatMessages() {
        try {
            return driver.findElements(By.xpath("//div[contains(@class,'message')] | //div[contains(@class,'bot-response')] | //div[contains(@class,'chat-bubble')]"));
        } catch (Exception e) {
            return new ArrayList<>();
        }
    }

    private String getLastChatMessageText() {
        List<WebElement> messages = getChatMessages();
        for (int i = messages.size() - 1; i >= 0; i--) {
            try {
                String text = messages.get(i).getText().trim();
                if (!text.isBlank()) {
                    return text;
                }
            } catch (StaleElementReferenceException ignored) {}
        }
        return "";
    }

    private String waitForChatbotResponse(String previousResponse) {
        WebDriverWait responseWait = new WebDriverWait(driver, RESPONSE_TIMEOUT);
        return responseWait.until(d -> {
            String currentResponse = getLastChatMessageText();
            if (currentResponse.isBlank()) return null;
            if (!currentResponse.equals(previousResponse)) return currentResponse;
            return null;
        });
    }

    private String sendQuestion(String question) {
        String previousResponse = getLastChatMessageText();
        WebElement chatInput = waitForChatInput();
        try {
            chatInput.click();
            chatInput.clear();
        } catch (Exception ignored) {}
        chatInput.sendKeys(question);
        chatInput.sendKeys(Keys.ENTER);
        return waitForChatbotResponse(previousResponse);
    }

    @Test
    public void runAutomationFramework() {
        List<TestRowData> testDataList = fetchExcelDataFromGitHub(FRAMEWORK_EXCEL_URL);
        Assert.assertFalse(testDataList.isEmpty(), "Failed to fetch Excel data or file is empty.");
        
        System.out.println("============================================================");
        System.out.println("Executing " + testDataList.size() + " test cases against target Chatbot.");
        System.out.println("============================================================");
        
        List<TestRowData> executedResults = new ArrayList<>();
        for (int i = 0; i < testDataList.size(); i++) {
            TestRowData rowData = testDataList.get(i);
            try {
                String chatbotResponse = sendQuestion(rowData.col3); // col3 is the question
                rowData.col10 = chatbotResponse; // Actual response written to column 10
                if (chatbotResponse == null || chatbotResponse.isBlank()) {
                    rowData.col12 = "FAIL";
                    rowData.col11 = "Chatbot returned an empty response.";
                } else if (chatbotResponse.toLowerCase().contains("error")) {
                    rowData.col12 = "FAIL";
                    rowData.col11 = "Chatbot returned an error message.";
                } else {
                    rowData.col12 = "PASS";
                    rowData.col11 = "Success.";
                }
            } catch (Exception e) {
                rowData.col10 = "EXCEPTION: " + e.getMessage();
                rowData.col12 = "FAIL";
                rowData.col11 = e.getMessage();
            }
            executedResults.add(rowData);
        }
        updateFrameworkExcel(executedResults);
    }

    private void updateFrameworkExcel(List<TestRowData> results) {
        try (InputStream is = openUrlStream(FRAMEWORK_EXCEL_URL);
             Workbook workbook = new XSSFWorkbook(is);
             FileOutputStream outputStream = new FileOutputStream(OUTPUT_EXCEL_FILE)) {
            for (TestRowData result : results) {
                Sheet sheet = workbook.getSheet(result.sheetName);
                if (sheet == null) continue;
                Row row = sheet.getRow(result.rowIndex);
                if (row == null) continue;
                
                // Write back outputs to columns 10, 11, and 12
                Cell answerCell = row.getCell(COL_CHATBOT_ANSWER);
                if (answerCell == null) answerCell = row.createCell(COL_CHATBOT_ANSWER);
                answerCell.setCellValue(result.col10);

                Cell reasonCell = row.getCell(COL_REASON);
                if (reasonCell == null) reasonCell = row.createCell(COL_REASON);
                reasonCell.setCellValue(result.col11);

                Cell statusCell = row.getCell(COL_STATUS);
                if (statusCell == null) statusCell = row.createCell(COL_STATUS);
                statusCell.setCellValue(result.col12);
            }
            workbook.write(outputStream);
            System.out.println("--> Execution results safely written up to Column 12 in output file.");
        } catch (Exception e) {
            throw new RuntimeException("Failed to update Excel: " + e.getMessage(), e);
        }
    }
}