package MiniProject;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;

public class ExcelReader {

	
	 @DataProvider(name = "DropdownData")
	    public Object[][] getData() throws IOException {
	        String filePath = "./data/Mysheet.xlsx";
	        FileInputStream fileInputStream = new FileInputStream("./data/Mysheet.xlsx");
	        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
	        XSSFSheet sheet = workbook.getSheetAt(0);

	        List<Object[]> data = new ArrayList<>();
	        Iterator<Row> rows = sheet.iterator();

	        // Skip header row if present
	        if (rows.hasNext()) rows.next();

	        // Iterate through each row and read cell data
	        while (rows.hasNext()) {
	            Row row = rows.next();
	            String[] rowData = new String[row.getLastCellNum()];

	            for (int i = 0; i < row.getLastCellNum(); i++) {
	                Cell cell = row.getCell(i);
	                rowData[i] = cell != null ? cell.toString() : "";
	            }
	            data.add(rowData);
	        }
	        workbook.close();
	        return data.toArray(new Object[0][0]);
	    }
}