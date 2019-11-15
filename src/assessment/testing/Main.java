package assessment.testing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Main {
	public static String vSearch;
	public static String[][] xData;
	public static int xlRows, xlCols;

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		Main m = new Main(); // object creation for the Main class
		m.execution(); // calling the function that holds the testing logic
	}

	//testing logic function
	public void execution() throws Exception {
		xlRead("C:\\Users\\vkhanna23\\Documents\\Training\\java\\testing\\assessment.testing\\YahooDDF.xls"); // reading data from xl sheet
		for (int i = 1; i < xlRows; i++) {
			if (xData[i][1].equals("Y")) {
				System.setProperty("webdriver.chrome.driver", "C:\\selenium jars\\chromedriver.exe"); // setting property for the web driver
				WebDriver driver = new ChromeDriver(); // initializing the web driver
				driver.get("https://www.yahoo.com"); // opening the yahoo.com page
				vSearch = xData[i][0]; // initialising the search text to be entered in the yahoo search bar
				driver.findElement(By.id("uh-search-box")).sendKeys(vSearch); // finding yahoo search bar and setting the value to search
				driver.findElement(By.id("uh-search-button")).click(); // finding yahoo search button and clicking it
				xData[i][2] = driver.getTitle(); // getting the title of the page
				driver.close();
			} else {
				xData[i][2] = "-";
			}
		}
		xlwrite("C:\\Users\\vkhanna23\\Documents\\Training\\java\\testing\\assessment.testing\\Result.xls", xData); // writing the value to the results xl file
	}

	// xl file reader function
	public static void xlRead(String sPath) throws Exception {
		File myFile = new File(sPath);
		FileInputStream myStream = new FileInputStream(myFile);
		HSSFWorkbook myworkbook = new HSSFWorkbook(myStream);
		HSSFSheet mySheet = myworkbook.getSheetAt(0);
		xlRows = mySheet.getLastRowNum() + 1;
		xlCols = mySheet.getRow(0).getLastCellNum();
		xData = new String[xlRows][xlCols];
		for (int i = 0; i < xlRows; i++) {
			HSSFRow row = mySheet.getRow(i);
			for (short j = 0; j < xlCols; j++) {
				HSSFCell cell = row.getCell(j);
				String value = cellToString(cell);
				xData[i][j] = value;
				System.out.print("-" + xData[i][j]);
			}
			System.out.println();
		}
	}

	// function to convert value in the cell to string
	public static String cellToString(HSSFCell cell) {
		int type = cell.getCellType();
		Object result;
		switch (type) {
		case HSSFCell.CELL_TYPE_NUMERIC:
			result = cell.getNumericCellValue();
			break;
		case HSSFCell.CELL_TYPE_STRING:
			result = cell.getStringCellValue();
			break;
		case HSSFCell.CELL_TYPE_FORMULA:
			throw new RuntimeException("We cannot evaluate formula");
		case HSSFCell.CELL_TYPE_BLANK:
			result = "-";
		case HSSFCell.CELL_TYPE_BOOLEAN:
			result = cell.getBooleanCellValue();
		case HSSFCell.CELL_TYPE_ERROR:
			result = "This cell has some error";
		default:
			throw new RuntimeException("We do not support this cell type");
		}
		return result.toString();

	}

	// xl file writer function
	public static void xlwrite(String xlpath1, String[][] xData) throws Exception {
		System.out.println("Inside XL Write");
		File myFile1 = new File(xlpath1);
		FileOutputStream fout = new FileOutputStream(myFile1);
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet mySheet1 = wb.createSheet("TestResults");
		for (int i = 0; i < xlRows; i++) {
			HSSFRow row1 = mySheet1.createRow(i);
			for (short j = 0; j < xlCols; j++) {
				HSSFCell cell1 = row1.createCell(j);
				cell1.setCellType(HSSFCell.CELL_TYPE_STRING);
				cell1.setCellValue(xData[i][j]);
			}
		}
		wb.write(fout);
		fout.flush();
		fout.close();
	}

}
