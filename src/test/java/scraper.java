import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.Collections;
import java.time.Duration;
import java.time.LocalDate;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Proxy;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import io.github.bonigarcia.wdm.WebDriverManager;

public class scraper {

	private static String URL = "https://www.nseindia.com/market-data/bonds-traded-in-capital-market";
	private static double DefaultCouponRate = 8;
	private static String sheetPath = "C:\\Users\\User\\Downloads\\";
	// main page locators
	private static By symbol = By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td/a)");
	private static By symbolCol = By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td/a)[1]");
	private static By seriesCol = By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[2]");
	private static By couponRateCol = By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[4]");
	private static By faceCol = By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[5]");
	private static By maturityDateCol = By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[12]");
	private static By noRecordsLbl = By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td[text()='No Records'])");
	private static By sortCouponRate = By.xpath("//a[@id='liveTCMTablecol3']");
	// new tab locators
	private static By siteCantBeReached = By.xpath("//span[text()='This site can’t be reached']");
	private static By askLbl = By.xpath("//th[contains(text(),'Ask')]/following::td[3]");
	private static By qtyLbl = By.xpath("//th[contains(text(),'Ask')]/following::td[4]");
	private static By ISINlbl = By.xpath("//div[@id='bondsAllSecurityTable']/table/tbody/tr/td[3]");
	private static By issueDescLbl = By.xpath("//div[@id='bondsAllSecurityTable']/table/tbody/tr/td[4]");

	private static double couponRateValue;
	private static String strCouponRate;
	private static String symbolValue;
	private static String seriesValue;
	private static String faceValue;
	private static String maturityDateValue;
	private static String askValue;
	private static String qtyValue;
	private static String ISINValue;
	private static String IssueDescValue;
	private static String[] excelData = new String[500];
	private static int lastRow = 1;
	private static String finalInterestValue;
	private static String f_timeRemain;
	private static double timeRemain;

	static WebDriver driver;
	static WebDriver driver2;

	public static void main(String[] args) throws Exception {
//		calculateTimeRem2();
//		calculateFinalRate2();
		System.out.println("STARTING...");
		WebDriverManager.chromedriver().setup();
//		System.setProperty("webdriver.chrome.driver", "C:\\WebDrivers\\chromedriver.exe");
//		// Set the proxy configuration
//		String proxyAddress = "52.67.10.183";
//		int proxyPort = 80;
//
//		// Create a Proxy object and set the HTTP and SSL proxies
//		Proxy proxy = new Proxy();
//		proxy.setHttpProxy(proxyAddress + ":" + proxyPort);
//		proxy.setSslProxy(proxyAddress + ":" + proxyPort);

		ChromeOptions options = new ChromeOptions();
//		options.setProxy(proxy);
		options.addArguments("--incognito");
		options.addArguments("--disable-infobars");
		options.addArguments("--start-maximized");
		options.addArguments("--disable-notifications");
//		options.addArguments("--disable-extensions");
//		options.addArguments("--disable-dev-shm-usage");
//		options.addArguments("--disable-impl-side-painting");
//		options.addArguments("--disable-gpu");
		options.addArguments("--no-sandbox");
		options.addArguments("--disable-setuid-sandbox");
		options.addArguments("--disable-dev-shm-using");
//		options.addArguments("--headless");
//		options.addArguments("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36");
//		options.addArguments("--disable-ipc-flooding-protection");
//		options.setExperimentalOption("excludeSwitches", Collections.singletonList("enable-automation"));
//		options.setExperimentalOption("useAutomationExtension", false);
//		options.addArguments("--user-agent=navneet");
		DesiredCapabilities capabilities = new DesiredCapabilities();
		capabilities.setCapability(ChromeOptions.CAPABILITY, options);
		options.merge(capabilities);
		driver = new ChromeDriver(options);
		driver.get(URL);
		Thread.sleep(4000);
//		String ipAddress = driver.findElement(By.tagName("body")).getText();
//		System.out.println("Your IP address: " + ipAddress); 

		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

		if (driver.findElements(noRecordsLbl).size() > 0) {
			System.out.println("NO RECORDS FOUND PAGE");
			driver.quit();
		} else {
			wait.until(ExpectedConditions.elementToBeClickable(symbol));
			driver.findElement(sortCouponRate).click();
			Thread.sleep(2000);
			driver.findElement(sortCouponRate).click();
			int totalRows = driver.findElements(symbol).size();
			int increment = 0;
//			WebDriver driver2 = new ChromeDriver(options);
			for (int i = 1; i <= totalRows; i++) {
				// check coupon rate > DefaultCouponRate
				String newURL;
				// for first set of records
				if (i == 1) {
					strCouponRate = driver
							.findElement(By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[" + (4) + "]")).getText()
							.trim();
					if (strCouponRate.equals("-")) {
						continue;
					} else {
						couponRateValue = Double.parseDouble(strCouponRate);
						System.out.println("CouponRate: " + couponRateValue);
						if (couponRateValue > DefaultCouponRate) {
							symbolValue = driver
									.findElement(By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td/a)[" + (1) + "]"))
									.getText().trim();
							seriesValue = driver
									.findElement(By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[" + (2) + "]"))
									.getText().trim();
							maturityDateValue = driver
									.findElement(By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[" + (12) + "]"))
									.getText().trim();
							faceValue = driver
									.findElement(By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[" + (5) + "]"))
									.getText().trim();
							// *******************NEW URL APPENDING USING DRIVER2************************
							newURL = "https://www.nseindia.com/get-quotes/bonds?symbol=" + symbolValue + "&series="
									+ seriesValue + "&maturityDate=" + maturityDateValue;
//							System.out.println("1:" + newURL);
							driver2 = new ChromeDriver(options);
							driver2.get(newURL);
							driver2.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(10));
							Thread.sleep(5000);
							getASKandQTY();
							driver2.close();
						} else {
							break;
						}
					}

				} // for remaining set of records
				else {
					increment += 12;
					strCouponRate = driver
							.findElement(By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[" + (4 + increment) + "]"))
							.getText().trim();
					if (strCouponRate.equals("-")) {
						continue;
					} else {
						couponRateValue = Double.parseDouble(strCouponRate);
						System.out.println("CouponRate: " + couponRateValue);
						if (couponRateValue > DefaultCouponRate) {
							symbolValue = driver
									.findElement(By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td/a)[" + i + "]"))
									.getText().trim();
							seriesValue = driver
									.findElement(By.xpath(
											"(//table[@id='liveTCMTable']/tbody/tr/td)[" + (2 + increment) + "]"))
									.getText().trim();
							maturityDateValue = driver
									.findElement(By.xpath(
											"(//table[@id='liveTCMTable']/tbody/tr/td)[" + (12 + increment) + "]"))
									.getText().trim();
							faceValue = driver
									.findElement(By.xpath(
											"(//table[@id='liveTCMTable']/tbody/tr/td)[" + (5 + increment) + "]"))
									.getText().trim();
							// *******************NEW URL APPENDING USING DRIVER2************************
							newURL = "https://www.nseindia.com/get-quotes/bonds?symbol=" + symbolValue + "&series="
									+ seriesValue + "&maturityDate=" + maturityDateValue;
//							System.out.println(i + ": " + newURL);
							driver2 = new ChromeDriver(options);
							driver2.get(newURL);
							driver2.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(10));
							Thread.sleep(5000);
							getASKandQTY();
							driver2.close();
						}
					}
				}
			}
			// copy files
			File sourceExcel = new File(sheetPath + "\\scrape_test.xlsx");
			File dstExcel = new File(sheetPath + "\\scrape_test" + outFileName() + ".xlsx");
			try {
				FileUtils.copyFile(sourceExcel, dstExcel);
			} catch (IOException e) {
				e.printStackTrace();
			}
			driver.quit();
			System.out.println("Successfully run");
		}
	}

	public static void calculateFinalRate() {
		String cleanedString = faceValue.replace(",", "");
		String cleanedString2 = askValue.replace(",", "");
		double f_faceValue = Double.parseDouble(cleanedString);
		double f_askValue = Double.parseDouble(cleanedString2);
		double out1, out2, out3, finalValue;
		calculateTimeRem();
		out1 = (f_faceValue - f_askValue);
		int roundedValue = (int) Math.ceil(timeRemain);
		out2 = ((couponRateValue * f_faceValue) / 100) * (roundedValue);
		out3 = (out1 + out2) / (timeRemain);
		finalValue = (out3 / f_askValue) * 100;
		System.out.println("Final: " + finalValue);
		DecimalFormat df = new DecimalFormat("0.00");
		// Format the double value
		finalInterestValue = df.format(finalValue);
		System.out.println("Final: " + finalInterestValue);
	}
	public static void calculateFinalRate2() {
//		String cleanedString = faceValue.replace(",", "");
//		String cleanedString2 = askValue.replace(",", "");
		double f_faceValue = 1000;//Double.parseDouble(cleanedString);
		double f_askValue = 908.35;//Double.parseDouble(cleanedString2);
		double out1, out2, out3, finalValue;
//		calculateTimeRem();
		out1 = (f_faceValue - f_askValue);
//		int roundedValue = (int) Math.ceil(timeRemain);
		out2 = ((8.65 * f_faceValue) / 100) * (4);
		out3 = (out1 + out2) / (3.78);
		finalValue = (out3 / f_askValue) * 100;
		System.out.println("Final: " + finalValue);
		DecimalFormat df = new DecimalFormat("0.00");
		// Format the double value
		finalInterestValue = df.format(finalValue);
		System.out.println("Final: " + finalInterestValue);
	}

	public static String calculateTimeRem() {
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MMM-yyyy");
		LocalDate date1 = LocalDate.parse(maturityDateValue, formatter);
		// Get today's date
		LocalDate today = LocalDate.now();
		// Calculate the difference in days
		long daysDifference = ChronoUnit.DAYS.between(today, date1);
		timeRemain = (double) daysDifference;
		timeRemain = timeRemain / 365;
		System.out.println("Days: " + timeRemain);
		DecimalFormat df = new DecimalFormat("0.00");
		// Format the double value
		f_timeRemain = df.format(timeRemain);
		return f_timeRemain;
	}
	
	public static String calculateTimeRem2() {
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MMM-yyyy");
		LocalDate date1 = LocalDate.parse("24-Jan-2028", formatter);
		// Get today's date
		LocalDate today = LocalDate.now();
		// Calculate the difference in days
		long daysDifference = ChronoUnit.DAYS.between(today, date1);
		timeRemain = (double) daysDifference;
		timeRemain = timeRemain / 365;
		System.out.println("Days: " + timeRemain);
		DecimalFormat df = new DecimalFormat("0.00");
		// Format the double value
		f_timeRemain = df.format(timeRemain);
		System.out.println("Days: " + f_timeRemain);
		return f_timeRemain;
	}

	public static void getASKandQTY() {
		if (driver2.findElements(siteCantBeReached).size() > 0) {
			System.out.println("NO RECORDS FOUND PAGE");
			driver2.close();
		} else if (!(driver2.findElements(askLbl).size() > 0)) {
			System.out.println("FIELDS NOT LOADING");
			excelData[0] = symbolValue;
			excelData[1] = seriesValue;
			excelData[2] = faceValue;
			excelData[3] = "NA";
			excelData[4] = "NA";
			excelData[5] = strCouponRate;
			excelData[6] = calculateTimeRem();
			excelData[7] = maturityDateValue;
			excelData[8] = "NA";
			excelData[9] = "NA";
			excelData[10] = "NA";
			writeExcelData(excelData);
		} else {
			askValue = driver2.findElement(askLbl).getText().trim();
			System.out.println(askValue);
			qtyValue = driver2.findElement(qtyLbl).getText().trim();
			System.out.println(qtyValue);
			ISINValue = driver2.findElement(ISINlbl).getText().trim();
			System.out.println(ISINValue);
			IssueDescValue = driver2.findElement(issueDescLbl).getText().trim();
			System.out.println(ISINValue);
			if (!askValue.equals("-")) {
				calculateFinalRate();
			}
			if (!askValue.equals("")) {
				excelData[0] = symbolValue;
				excelData[1] = seriesValue;
				excelData[2] = faceValue;
				excelData[3] = askValue;
				excelData[4] = qtyValue;
				excelData[5] = strCouponRate;
				excelData[6] = calculateTimeRem();
				excelData[7] = maturityDateValue;
				excelData[8] = ISINValue;
				excelData[9] = IssueDescValue;
				if (askValue.equals("-")) {
					excelData[3] = "0";
					excelData[4] = "0";
					excelData[10] = "0";
				} else {
					excelData[10] = (finalInterestValue);
				}
				writeExcelData(excelData);
			}
		}
	}

	public static String outFileName() {
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd-HH-mm-ss");
		LocalDateTime now = LocalDateTime.now();
		System.out.println(dtf.format(now));
		return dtf.format(now).toString();
	}

	public static void writeExcelData(String[] symbol) {

		XSSFWorkbook workbook = null;
		try {
			FileInputStream file = new FileInputStream(new File(sheetPath + "\\scrape_test.xlsx")); // "C:\Users\User\Downloads\scrape_test.xlsx"
			workbook = new XSSFWorkbook(file);
			XSSFSheet wSheet = workbook.getSheet("Sheet1");
			int lastColNum = 11;
			XSSFCell cell;
			XSSFRow row;

			// Clear all rows from row 1 on initial run
			if (lastRow == 1) {
				int lastRowNum1 = wSheet.getLastRowNum();
				for (int i = 1; i <= lastRowNum1; i++) {
					Row row1 = wSheet.getRow(i);
					if (row1 != null) {
						wSheet.removeRow(row1);
					}
				}
			}
			Thread.sleep(3000);
//			row = wSheet.createRow(wSheet.getLastRowNum() + 1);
			row = wSheet.createRow(lastRow);
			for (int j = 0; j < lastColNum; j++) {
				cell = row.createCell(j);
				cell.setCellValue(symbol[j]);
			}
			lastRow += 1;
			file.close();

		} catch (Exception e) {
			e.printStackTrace();
		}
		try {
			FileOutputStream out = new FileOutputStream(new File(sheetPath + "\\scrape_test.xlsx"));
			workbook.write(out);
			workbook.close();
			out.close();
			System.out.println("Output generated successfully");
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}
