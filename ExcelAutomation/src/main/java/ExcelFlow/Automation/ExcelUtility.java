package ExcelFlow.Automation;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ExcelUtility {

	public void getRowCount() {

		try {

			FileInputStream is = new FileInputStream(
					"C:\\Users\\HP\\eclipse-workspace\\ExcelAutomation\\src\\main\\java\\ExcelFile\\ExcelSheet.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(is);
			XSSFSheet sheet = workbook.getSheet("Sheet1");
			int rowCount = sheet.getPhysicalNumberOfRows();
			System.out.println("Row count= " + rowCount);

		} catch (IOException exp) {
			// TODO Auto-generated catch block

			String Message = exp.getMessage();
			System.out.println(Message);
			exp.getCause();
			exp.printStackTrace();
		}

	}

	public void getCellData(int rowNum, int ColumnNum) throws IOException {

		FileInputStream is = new FileInputStream(
				"C:\\Users\\HP\\eclipse-workspace\\ExcelAutomation\\src\\main\\java\\ExcelFile\\ExcelSheet.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(is);
		XSSFSheet sheet = workbook.getSheet("Sheet1");

		String username = sheet.getRow(rowNum).getCell(ColumnNum).getStringCellValue();
		System.out.println("Data=" + username);

	}

	public static void main(String[] args) throws IOException {
		ExcelUtility executeExcel = new ExcelUtility();
		// executeExcel.getRowCount();
		executeExcel.getCellData(0, 0);

	}
}
