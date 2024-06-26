import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.Collections;
import java.util.Locale;
import java.util.concurrent.TimeUnit;
import java.time.Duration;
import java.time.LocalDate;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import io.github.bonigarcia.wdm.WebDriverManager;
//import io.github.bonigarcia.wdm.WebDriverManager;

public class scraper {

	private static String NSEURL = "https://www.nseindia.com/market-data/bonds-traded-in-capital-market";
	private static String iciciDirectURL = "https://www.icicidirect.com/fd-and-bonds";
	private static double DefaultCouponRate = 8;
	private static int lastColNum = 13;
	private static String sheetPath = System.getProperty("user.dir") + "\\scrape_test.xlsx"; // "C:\\Users\\User\\Downloads\\";
	private static String sheetOutPath = System.getProperty("user.dir") + "\\scrape_test_" + outFileName() + ".xlsx";
	// main page locators NSE
	private static By symbol = By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td/a)");
//	private static By symbolCol = By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td/a)[1]");
//	private static By seriesCol = By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[2]");
//	private static By couponRateCol = By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[4]");
//	private static By faceCol = By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[5]");
//	private static By maturityDateCol = By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[12]");
	private static By noRecordsLbl = By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td[text()='No Records'])");
	private static By sortCouponRate = By.xpath("//a[@id='liveTCMTablecol3']");
	private static By notLoadingLbl = By.xpath("(//div[@class='loader-wrp'])[2]");
	// new tab locators NSE bond
	private static By siteCantBeReached = By.xpath("//span[text()='This site can’t be reached']");
	private static By askLbl = By.xpath("//th[contains(text(),'Ask')]/following::td[3]");
	private static By qtyLbl = By.xpath("//th[contains(text(),'Ask')]/following::td[4]");
	private static By ISINlbl = By.xpath("//div[@id='bondsAllSecurityTable']/table/tbody/tr/td[3]");
	private static By issueDescLbl = By.xpath("//div[@id='bondsAllSecurityTable']/table/tbody/tr/td[4]");
	// ICICI page locators
	private static By frequencyLbl = By.xpath("//h4[text()='Coupon Frequency ']/following::h5[1]");
	private static By yieldICICILbl = By.xpath("//h4[text()='Yield ']/following::h5[1]");
	private static By searchTxt = By.xpath("//input[@id='searchStock']");
//	private static By clickISIN = By.xpath("//span[contains(text(),'INE530B08102')]");

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
	private static String finalInterestYieldValue;
	private static String finalYieldPerAnnumValue;
	private static String f_timeRemain;
	private static double timeRemain;

	private static String freqValue;
	private static String yieldICICIValue;
	private static String netGainValue;

	static WebDriver driver;
	static WebDriver driver2;
	static WebDriver driver3;
	static ChromeOptions options = new ChromeOptions();
//	static EdgeOptions options = new EdgeOptions();

	public static void main(String[] args) throws Exception {
		try {
			System.out.println("STARTING...");
			WebDriverManager.chromedriver().setup();
//			WebDriverManager.edgedriver().setup();
//			System.setProperty("webdriver.chrome.driver", "C:\\WebDrivers\\chromedriver.exe");

			options.addArguments("--incognito");
			options.addArguments("--disable-infobars");
			options.addArguments("--window-size=1366,768");
			options.addArguments("--start-maximized");
			options.addArguments("--disable-notifications");
			options.addArguments("--disable-extensions");
//			options.addArguments("--disable-dev-shm-usage");
//			options.addArguments("--disable-impl-side-painting");
//			options.addArguments("--disable-gpu");
//			options.addArguments("--no-sandbox");
//			options.addArguments("--disable-setuid-sandbox");
//			options.addArguments("--disable-dev-shm-using");
			options.addArguments("--disable-blink-features=AutomationControlled");// ***************EUREKA**************
			options.addArguments("--headless");
			options.addArguments(
					"user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36");
			options.addArguments("--disable-ipc-flooding-protection");
			options.setExperimentalOption("excludeSwitches", Collections.singletonList("enable-automation"));
			options.setExperimentalOption("useAutomationExtension", false);
			DesiredCapabilities capabilities = new DesiredCapabilities();
			capabilities.setCapability(ChromeOptions.CAPABILITY, options);
			options.merge(capabilities);
			driver = new ChromeDriver(options);
			driver.get(NSEURL);
			long startTime = System.nanoTime();
			long endTime;
			long elapsedTimeInMillis;
			Thread.sleep(4000);

			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));

			if (driver.findElements(noRecordsLbl).size() > 0) {
				System.out.println("NO RECORDS FOUND PAGE");
				driver.quit();
			} else if (driver.findElements(notLoadingLbl).size() > 0) {
				System.out.println("RECORDS NOT LOADING, LOADING WHEEL DISPLAYED");
				driver.quit();
			} else {

				// NSE BOND site now
				wait.until(ExpectedConditions.elementToBeClickable(symbol));
				driver.findElement(sortCouponRate).click();
				Thread.sleep(2000);
				driver.findElement(sortCouponRate).click();
				int totalRows = driver.findElements(symbol).size();
				int increment = 0;
				// open ICICI direct site for frequency
				driver3 = new ChromeDriver(options);
				driver3.get(iciciDirectURL);

				for (int i = 1; i <= totalRows; i++) {
					endTime = System.nanoTime();
					elapsedTimeInMillis = TimeUnit.NANOSECONDS.toMinutes(endTime - startTime);
					System.out.println("Time elapsed: " + elapsedTimeInMillis + " mins");
					// check coupon rate > DefaultCouponRate
					String bondNSEURL;
					// for first set of records
					if (i == 1) {
						strCouponRate = driver
								.findElement(By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[" + (4) + "]"))
								.getText().trim();
						if (strCouponRate.equals("-")) {
							continue;
						} else {
							couponRateValue = Double.parseDouble(strCouponRate);
							System.out.println(
									"*******************************************************************************************");
							System.out.println("CouponRate: " + couponRateValue);
							if (couponRateValue > DefaultCouponRate) {
								symbolValue = driver
										.findElement(
												By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td/a)[" + (1) + "]"))
										.getText().trim();
								seriesValue = driver
										.findElement(By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[" + (2) + "]"))
										.getText().trim();
								maturityDateValue = driver
										.findElement(
												By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[" + (12) + "]"))
										.getText().trim();
								faceValue = driver
										.findElement(By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[" + (5) + "]"))
										.getText().trim();
								// *******************NEW URL APPENDING USING DRIVER2************************
								bondNSEURL = "https://www.nseindia.com/get-quotes/bonds?symbol=" + symbolValue
										+ "&series=" + seriesValue + "&maturityDate=" + maturityDateValue;
								System.out.println(symbolValue + "__" + seriesValue);
								driver2 = new ChromeDriver(options);
								driver2.get(bondNSEURL);
								Thread.sleep(3000);
								getASKandQTY();
								driver2.quit();
							}
//							else {
//								break;
//							}
						}

					} // for remaining set of records
					else {
						increment += 12;
						strCouponRate = driver
								.findElement(
										By.xpath("(//table[@id='liveTCMTable']/tbody/tr/td)[" + (4 + increment) + "]"))
								.getText().trim();
						if (strCouponRate.equals("-")) {
							continue;
						} else {
							couponRateValue = Double.parseDouble(strCouponRate);
							System.out.println(
									"*******************************************************************************************");
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
								bondNSEURL = "https://www.nseindia.com/get-quotes/bonds?symbol=" + symbolValue
										+ "&series=" + seriesValue + "&maturityDate=" + maturityDateValue;
								System.out.println(symbolValue + "__" + seriesValue);
								driver2 = new ChromeDriver(options);
								driver2.get(bondNSEURL);
								Thread.sleep(3000);
								getASKandQTY();
								driver2.quit();
							} else {
								break;
							}
						}
					}
				}
				// copy files
				File sourceExcel = new File(sheetPath);
				File dstExcel = new File(sheetOutPath);
				try {
					FileUtils.copyFile(sourceExcel, dstExcel);
				} catch (IOException e) {
					e.printStackTrace();
				}
				endTime = System.nanoTime();
				elapsedTimeInMillis = TimeUnit.NANOSECONDS.toMinutes(endTime - startTime);
				System.out.println("Time taken to run the script: " + elapsedTimeInMillis + " mins");
				driver.quit();
				driver3.quit();
				System.out.println("Successfully run, O/P file generated");
			}
		} catch (Exception e) {
			// Handle exceptions
			e.printStackTrace();
			// Add any specific exception handling here
		} finally {
			// Ensure WebDriver instance is closed
			if (driver != null) {
				driver.quit();
			}
			if (driver2 != null) {
				driver2.quit();
			}
			if (driver3 != null) {
				driver3.quit();
			}

		}
	}

	public static void checkFrequency(String ISIN) throws Exception {
		WebDriverWait wait2 = new WebDriverWait(driver3, Duration.ofSeconds(20));
		driver3.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(20));
		driver3.findElement(searchTxt).clear();
		Thread.sleep(1000);
		driver3.findElement(searchTxt).sendKeys(ISIN);
		driver3.findElement(searchTxt).sendKeys(Keys.ENTER);
		Thread.sleep(3000);
		freqValue = "0";
		if (driver3.findElements(By.xpath("//label[not(@style='display:none')][text()='No records found']")).size() > 0
				&& driver3.findElements(By.xpath("//label[not(@style='display: none;')][text()='No records found']"))
						.size() > 0
				&& driver3.findElements(By.xpath("//label[(@style)][text()='No records found']")).size() > 0) {
			System.out.println("**************BOND NOT FOUND****************");
			freqValue = "0";
//			yieldICICIValue = "0";
		} else {
			try {
				wait2.until(
						ExpectedConditions.elementToBeClickable(By.xpath("//span[contains(text(),'" + ISIN + "')]")));
				driver3.findElement(By.xpath("//span[contains(text(),'" + ISIN + "')]")).click();
				Thread.sleep(4000);
				freqValue = driver3.findElement(frequencyLbl).getText().trim();
				System.out.println("Freq: " + freqValue);
//				yieldICICIValue = driver3.findElement(yieldICICILbl).getText().trim();
//				System.out.println("YieldICICI: " + yieldICICIValue);
			} catch (TimeoutException e) {
				System.out.println("Timeout exception occurred: " + e.getMessage());
			} catch (NoSuchElementException e) {
				System.out.println("Element not found: " + e.getMessage());
			}
		}
	}

	public static void calculateFinalRate() {
		double out1, out2, out3, finalValue;
		calculateTimeRem();
		out1 = (convertCommaToDouble(faceValue) - convertCommaToDouble(askValue));
		int roundedValue = (int) Math.ceil(timeRemain);

		// for calculating SI & final yield PA
		int time;
		double rate;
		if (freqValue.equalsIgnoreCase("Yearly") || freqValue.equalsIgnoreCase("Annual")) {
			time = (int) Math.ceil(timeRemain * 1);
			rate = couponRateValue / 1;
			calculateFinalYieldPA(time, rate);
		} else if (freqValue.equalsIgnoreCase("Monthly")) {
			time = (int) Math.ceil(timeRemain * 12);
			rate = couponRateValue / 12;
			calculateFinalYieldPA(time, rate);
		} else if (freqValue.equalsIgnoreCase("Quarterly")) {
			time = (int) Math.ceil(timeRemain * 4);
			rate = couponRateValue / 4;
			calculateFinalYieldPA(time, rate);
		} else {
			finalYieldPerAnnumValue = "NA";
			System.out.println("NO DATA");
		}

		// for calculating final yield interest
		out2 = ((couponRateValue * convertCommaToDouble(faceValue)) / 100) * (roundedValue);
		out3 = (out1 + out2) / (timeRemain);
		finalValue = (out3 / convertCommaToDouble(askValue)) * 100;
		finalInterestYieldValue = convertToStringAndTwoDecimal(finalValue);
		System.out.println("yield: " + finalInterestYieldValue);
	}

	public static void calculateFinalYieldPA(int time, double rate) {
		double SI, yieldOnInvest, discountOrPremium, netGain;
		SI = (convertCommaToDouble(faceValue) * rate * time) / 100;
		System.out.println("time: " + time + " rate: " + rate + " SI: " + SI);
//		yieldOnInvest = SI / convertCommaToDouble(askValue);
		discountOrPremium = convertCommaToDouble(faceValue) - convertCommaToDouble(askValue);
		netGain = SI + discountOrPremium;
		System.out.println("Discount/premium: " + discountOrPremium + " net gain: " + netGain);
		netGainValue = convertToStringAndTwoDecimal(netGain);
		yieldOnInvest = netGain / convertCommaToDouble(askValue);
		System.out.println("yield on invest: " + yieldOnInvest);
		finalYieldPerAnnumValue = convertToStringAndTwoDecimal(
				(yieldOnInvest / Double.parseDouble(f_timeRemain)) * 100);
		System.out.println("Final yield p.a.: " + finalYieldPerAnnumValue);
	}

	public static double convertCommaToDouble(String strValue) {
		String cleanedString = strValue.replace(",", "");
		double doubleValue = Double.parseDouble(cleanedString);
		return doubleValue;
	}

	public static double convertToDouble(String strValue) {
		double doubleValue = Double.parseDouble(strValue);
		return doubleValue;
	}

	public static String convertToStringAndTwoDecimal(double dbValue) {
		DecimalFormat df = new DecimalFormat("0.00");
		String strValue = df.format(dbValue);
		return strValue;
	}

	public static String calculateTimeRem() {
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MMM-yyyy", Locale.ENGLISH);
		LocalDate date1 = LocalDate.parse(maturityDateValue, formatter);
		// Get today's date
		LocalDate today = LocalDate.now();
		// Calculate the difference in days
		long daysDifference = ChronoUnit.DAYS.between(today, date1);
		timeRemain = (double) daysDifference;
		timeRemain = timeRemain / 365;
		f_timeRemain = convertToStringAndTwoDecimal(timeRemain);
		System.out.println("Time period remaining: " + f_timeRemain);
		return f_timeRemain;
	}

	public static void getASKandQTY() throws Exception {
		if (driver2.findElements(siteCantBeReached).size() > 0) {
			System.out.println("NO RECORDS FOUND PAGE");
			driver2.quit();
		} else if (!(driver2.findElements(askLbl).size() > 0)) {
			System.out.println("FIELDS NOT LOADING");
			excelData[0] = symbolValue;
			excelData[1] = seriesValue;
			excelData[2] = faceValue;
			excelData[3] = "NA";
			excelData[4] = "NA";
			excelData[5] = "NA";
			excelData[6] = strCouponRate;
			excelData[7] = calculateTimeRem();
			excelData[8] = maturityDateValue;
			excelData[9] = "NA";
			excelData[10] = "NA";
			excelData[11] = "NA";
			excelData[12] = "NA";
			excelData[13] = "NA";
			System.out.println("BEEP BEEP");
			writeExcelData(excelData);
		} else {
			askValue = driver2.findElement(askLbl).getText().trim();
			System.out.println("Ask: " + askValue);
			qtyValue = driver2.findElement(qtyLbl).getText().trim();
			System.out.println("Qty: " + qtyValue);
			ISINValue = driver2.findElement(ISINlbl).getText().trim();
			System.out.println("ISIN: " + ISINValue);
			IssueDescValue = driver2.findElement(issueDescLbl).getText().trim();
			System.out.println("ISIN desc:" + IssueDescValue);

			if (!askValue.equals("-") && (!askValue.equals(""))) {
				// check frequency on ICICI direct site
				checkFrequency(ISINValue);
				calculateFinalRate();
				excelData[0] = symbolValue;
				excelData[1] = seriesValue;
				excelData[2] = faceValue;
				excelData[3] = askValue;
				excelData[4] = qtyValue;

				excelData[6] = strCouponRate;
				excelData[7] = f_timeRemain;
				excelData[8] = maturityDateValue;
				excelData[9] = ISINValue;
				excelData[10] = IssueDescValue;
				excelData[5] = convertToStringAndTwoDecimal(
						convertCommaToDouble(askValue) * convertCommaToDouble(qtyValue));
				excelData[11] = freqValue;
				excelData[12] = netGainValue;
//				excelData[13] = (finalInterestYieldValue);
				excelData[13] = (finalYieldPerAnnumValue);
				writeExcelData(excelData);
			}
//			if (!askValue.equals("")) {
//				excelData[0] = symbolValue;
//				excelData[1] = seriesValue;
//				excelData[2] = faceValue;
//				excelData[3] = askValue;
//				excelData[4] = qtyValue;
//				excelData[5] = "x";
//				excelData[6] = strCouponRate;
//				excelData[7] = f_timeRemain;
//				excelData[8] = maturityDateValue;
//				excelData[9] = ISINValue;
//				excelData[10] = IssueDescValue;
//				if (askValue.equals("-")) {
//					excelData[3] = "0";
//					excelData[4] = "0";
//					excelData[5] = "0";
//					excelData[11] = "0";
//					excelData[12] = "0";
//					excelData[13] = "0";
////					excelData[14] = "0";
//				} else {
//					excelData[5] = convertToStringAndTwoDecimal(
//							convertCommaToDouble(askValue) * convertCommaToDouble(qtyValue));
//					excelData[11] = freqValue;
//					excelData[12] = netGainValue;
////					excelData[13] = (finalInterestYieldValue);
//					excelData[13] = (finalYieldPerAnnumValue);
//				}
//				writeExcelData(excelData);
//			}
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
			FileInputStream file = new FileInputStream(new File(sheetPath)); // "C:\Users\User\Downloads\scrape_test.xlsx"
			workbook = new XSSFWorkbook(file);
			XSSFSheet wSheet = workbook.getSheet("Sheet1");
//			int lastColNum = 13;
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
//			row = wSheet.createRow(wSheet.getLastRowNum() + 1);
			row = wSheet.createRow(lastRow);
			for (int j = 0; j <= lastColNum; j++) {
				cell = row.createCell(j);
				cell.setCellValue(symbol[j]);
			}
			lastRow += 1;
			file.close();

		} catch (Exception e) {
			e.printStackTrace();
		}
		try {
			FileOutputStream out = new FileOutputStream(new File(sheetPath));
			workbook.write(out);
			workbook.close();
			out.close();
			System.out.println("Output generated successfully");
			System.out.println(
					"*******************************************************************************************");
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}
