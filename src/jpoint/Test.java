package jpoint;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;

public class Test {
	File F1;
	FileInputStream FIS;
	XSSFWorkbook WB;
	XSSFSheet WS;
	FileOutputStream FOS;
	int i;
	WebDriver driver;
	List<WebElement> ML;
	List<WebElement> LT;
	List<WebElement> PT;
	List<WebElement> historylist;
	List<WebElement> feature;
	String history;
	

	public void launchurl() {
		driver = new FirefoxDriver();
		driver.get("https://www.javatpoint.com/");
	}

	public void PopularTutorials() throws Throwable {
		PT = driver.findElements(By.xpath(".//*[@id='city']/table/tbody/tr/td/fieldset[2]/div/a/div/span"));

	}

	public void menu() {
		ML = driver.findElements(By.xpath(".//*[@id='link']/div/ul/li/a"));

	}

	public void Latest5Tutorials() throws Throwable {
		LT = driver.findElements(By.xpath(".//*[@id='city']/table/tbody/tr/td/fieldset[1]/div/a/div/span"));

	}

	public void switchwindow() {
		WebElement java = driver.findElement(By.xpath(".//*[@id='link']/div/ul/li[2]/a"));
		Actions ACT = new Actions(driver);
		ACT.contextClick(java).sendKeys(Keys.ARROW_DOWN).sendKeys(Keys.ARROW_DOWN).sendKeys(Keys.ENTER).perform();
		String Parentwindow = driver.getWindowHandle();
		Set<String> S1 = driver.getWindowHandles();
		Iterator<String> I1 = S1.iterator();
		while (I1.hasNext()) {
			String childwindow = I1.next();
			if (!Parentwindow.equalsIgnoreCase(childwindow)) {
				driver.switchTo().window(childwindow);
				System.out.println(driver.findElement(By.xpath(".//*[@id='city']/table/tbody/tr/td/h1")).getText());
				driver.findElement(By.xpath(".//*[@id='menu']/div[2]/a[2]")).click();
				history = driver.findElement(By.xpath(".//*[@id='city']/table/tbody/tr/td/h3")).getText();
				System.out.println(history);
				historylist = driver.findElements(By.xpath(".//*[@id='city']/table/tbody/tr/td/div[4]/ol/li"));
			}
		}
	}

	public void excell() throws Throwable {
		F1 = new File("C:\\Users\\PrathimaKandela\\Documents\\jpoint.xlsx");
		FIS = new FileInputStream(F1);
		WB = new XSSFWorkbook(FIS);
		// WB= new XSSFWorkbook(FIS);
		WS = WB.getSheet("Sheet1");
		for (int k = 0; k <= PT.size() - 1; k++) {
			String PTlist = PT.get(k).getText();
			System.out.println("Popular Tutorials=" + PTlist);
			WS = WB.getSheet("Sheet1");
			WS.createRow(0).createCell(0).setCellValue("Popular Tutorials");
			WS.createRow(k + 1).createCell(1).setCellValue(PTlist);
		}
		for (i = 0; i <= ML.size() - 1; i++) {
			String MLlist = ML.get(i).getText();
			System.out.println("MENU=" + MLlist);
			WS.getRow(0).createCell(1).setCellValue("MENU");

			WS.getRow(i + 1).createCell(1).setCellValue(MLlist);

		}
		for (int j = 0; j <= LT.size() - 1; j++) {
			String LTlist = LT.get(j).getText();
			System.out.println("Latest 5 Tutorials=" + LTlist);
			WS = WB.getSheet("Sheet1");
			WS.getRow(0).createCell(2).setCellValue("Latest 5 Tutorials");
			WS.getRow(j + 1).createCell(2).setCellValue(LTlist);

		}
		for (int l = 0; l <= historylist.size() - 1; l++) {
			String HL = historylist.get(l).getText();
			System.out.println("historylist=" + HL);
			WS.getRow(0).createCell(3).setCellValue("historylist");
			WS.getRow(l + 1).createCell(3).setCellValue(HL);
		}
		FOS = new FileOutputStream(F1);
		WB.write(FOS);
		WB.close();
		FOS.close();
		driver.quit();
	}

	public static void main(String[] args) throws Throwable {

		Test t = new Test();
		t.launchurl();
		t.PopularTutorials();
		t.menu();
		t.Latest5Tutorials();
		t.switchwindow();
		t.excell();


	}

}
