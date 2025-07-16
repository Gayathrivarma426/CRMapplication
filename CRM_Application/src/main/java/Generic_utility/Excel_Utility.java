package Generic_utility;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Excel_Utility {

	

	/**
	 * This method is used to read data from excel file
	 * @param SheetName
	 * @param rowNum
	 * @param cellNum
	 * @return
	 * @throws Throwable
	 * @author gayathri
	 */
	public String getExcelData(String Sheet,int Row,int Cell) throws Throwable {
	
		FileInputStream fes1 = new FileInputStream("C:\\Users\\sures\\Downloads\\VtigerData.xlsx");
	     
	    Workbook book = WorkbookFactory.create(fes1);
	
	     Sheet sheet = book.getSheet(Sheet);
	     
	    Row row1 = sheet.getRow(Row);
	    Cell cell1 = row1.getCell(Cell);
	    String value = cell1.getStringCellValue();
	    return  value;
	}
	public String getExcelDataUsingFormatter(String sheet1,int Row1,int Cell1) throws Throwable
	{
		FileInputStream fes1 = new FileInputStream("C:\\Users\\sures\\Downloads\\VtigerData.xlsx");
	     
	    Workbook book = WorkbookFactory.create(fes1);
	
	     Sheet sheet = book.getSheet(sheet1);
	     
	    Row row1 = sheet.getRow(Row1);
	    Cell cell1 = row1.getCell(Cell1);
	    DataFormatter format = new DataFormatter();
		String Exceldata = format.formatCellValue(cell1);
		return Exceldata;
		
	}
	
	public  Object[][] readDataUsingDAtaProvider(String SheetName) throws Throwable
	{
		
		FileInputStream fes = new FileInputStream("C:\\Users\\sures\\Downloads\\VtigerData.xlsx");

		// step2:- keeps the workbook/Excel file in read mode
		Workbook book = WorkbookFactory.create(fes);

		// step3:- Navigates into mentioned sheetname
		Sheet sheet = book.getSheet(SheetName);
		
		int lastRow = sheet.getLastRowNum()+1;
		
		int lastCell = sheet.getRow(0).getLastCellNum();
		
		Object[][] objArr = new Object[lastRow][lastCell];
		
		for(int i=0;i<lastRow;i++)
		{
			for (int j = 0; j < lastCell; j++)
			{
				objArr[i][j]=sheet.getRow(i).getCell(j).toString();
			}
		}
		
		return objArr;
		
	}
}
