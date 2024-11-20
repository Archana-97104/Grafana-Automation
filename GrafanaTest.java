package MiniProject;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;

public class GrafanaTest {
	private XSSFWorkbook outputWorkbook;
	private XSSFSheet outputSheet;
	private String outputFilePath;
	private WebDriver driver;

	@BeforeSuite
	public void setUp() throws IOException {

		outputFilePath = "C:\\Users\\604550803\\OneDrive - NBCUniversal\\My Documents\\Desktop\\JAVA-IN\\Eclipse\\Eclipse\\eclipse\\Datas\\Automationproject\\Data\\console_output.xlsx";

		outputWorkbook = new XSSFWorkbook();
		outputSheet = outputWorkbook.createSheet("Console Output");
		XSSFRow headerRow = outputSheet.createRow(0);
		headerRow.createCell(0).setCellValue("Environment");
		headerRow.createCell(1).setCellValue("Region");
		headerRow.createCell(2).setCellValue("Query");
		headerRow.createCell(3).setCellValue("Session");
		headerRow.createCell(4).setCellValue("Result Text");
	}

	@BeforeMethod
	public void setUp1() {
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	@Test(dataProvider = "DropdownData", dataProviderClass = ExcelReader.class)
	public void runGrafana(String envInput, String regionInput, String queryInput, String sessionInput)
			throws IOException {
		// to open chrome

		driver.get("https://d3q7rt0kr5fynf.cloudfront.net/ad-tools/mediatailor-logs-query-builder.html");
		// to select dropdown
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
		WebElement dropdown1 = driver.findElement(By.id("envInput"));
		Select option1 = new Select(dropdown1);
		option1.selectByVisibleText(envInput);
		WebElement dropdown2 = driver.findElement(By.id("regionInput"));
		Select option2 = new Select(dropdown2);
		option2.selectByVisibleText(regionInput);
		WebElement dropdown3 = driver.findElement(By.id("queryInput"));
		Select option3 = new Select(dropdown3);
		option3.selectByVisibleText(queryInput);
		driver.findElement(By.id("sessionInput")).sendKeys(sessionInput);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		driver.findElement(By.id("button-mt")).click();
		Set<String> windowHandles = driver.getWindowHandles();
		List<String> allWindows = new ArrayList<String>(windowHandles);
		driver.switchTo().window(allWindows.get(1));
		driver.findElement(By.xpath("//span[@class='css-1ueg5w']//span[1]")).click();
		driver.findElement(By.xpath("//label[text()='Last 6 hours']")).click();
		driver.findElement(By.xpath("//button[@class= 'css-5se5b3 css-1wx8bl8-positionRelative']")).click();
		String text = driver.findElement(By.xpath("//button[@class= 'css-5se5b3 css-1wx8bl8-positionRelative']"))
				.getText();

		// to write results in excel
		int lastRow = outputSheet.getLastRowNum() + 1;
		XSSFRow row = outputSheet.createRow(lastRow);
		row.createCell(0).setCellValue(envInput);
		row.createCell(1).setCellValue(regionInput);
		row.createCell(2).setCellValue(queryInput);
		row.createCell(3).setCellValue(sessionInput);
		row.createCell(4).setCellValue(text);

		System.out.println("Result written to row: " + lastRow);
	}

	@AfterMethod
	public void tearDown() {
		if (driver != null) {

		}
	}

	@AfterSuite
	public void tearDownSuite() throws IOException {
		if (outputWorkbook != null) {
			try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
				outputWorkbook.write(fos);
			}
			outputWorkbook.close();
		}
	}

}
