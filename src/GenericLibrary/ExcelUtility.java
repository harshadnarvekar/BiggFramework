package GenericLibrary;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class ExcelUtility {

	/** 
	 * Desc - This returns a handle of the Excelsheet
	 * @param fileName - Path and the file we need to excess (File)
	 * @param sheetName - Sheet name of the file (String)
	 * @return - XSSFSheet handle
	 * @throws Exception
	 */
	public static XSSFSheet GetSheetHandle(File fileName, String sheetName) throws Exception
	{
		try
		{
			FileInputStream mcExcelFileHndl = new FileInputStream(fileName);
			@SuppressWarnings("resource")
			XSSFWorkbook mcWorkbookHndl = new XSSFWorkbook(mcExcelFileHndl); 
			XSSFSheet mcSheet = mcWorkbookHndl.getSheet(sheetName);
			return mcSheet;
		}
		catch(Exception e)
		{
			System.out.println("Exception caught in "+ e);
			return null;
		}
	}
	/**
	 * Desc - Returns the number of rows present in the excelsheet 
	 * @param sheetHandle - Sheet handle of the excel file of which we need to find the count (XSSFSheet handle) 
	 * @return - the count in Integer
	 */
	public static Integer getRowCount(XSSFSheet sheetHandle)
	{
		Integer rowCount = null;
		rowCount = sheetHandle.getLastRowNum();
		return rowCount;
	}
	
	
	public static Integer getStartingExcecutionRowOfMC(XSSFSheet sheetHandle) throws Exception
	{
		int rowNo = 0;
		
		try
		{
			int mcRowCount,mcRowCounter;
			int newRow = sheetHandle.getPhysicalNumberOfRows();
//			System.out.println("physrical no of rows are"+ newRow);
			mcRowCount = ExcelUtility.getRowCount(sheetHandle);
			System.out.println("physrical no of rows are"+ mcRowCount);

			
			
			Boolean bStartPosFound, bYPosFound = false;

			// check the path file name and start looping it
			for(mcRowCounter=1;mcRowCounter<mcRowCount+1;mcRowCounter++)
			{
				XSSFRow eachRow = null;
				XSSFCell cellValueYPos, cellValueStartPos = null;
				eachRow = sheetHandle.getRow(mcRowCounter);
				cellValueYPos = eachRow.getCell(5);
				cellValueStartPos = eachRow.getCell(6);
				//System.out.println("cell value of Y at row no.-"+mcRowCounter+"value is-"+cellValueYPos.toString()+"- and the other start-"+cellValueStartPos.toString() );
				if(cellValueYPos.toString().equalsIgnoreCase("Y") && cellValueStartPos.toString().equalsIgnoreCase("START"))
				{
					
					System.out.println("inside FIRST if");
					 bStartPosFound = true; // global variable
						 bYPosFound = true; // global variable
						System.out.println("Yahoo here we start");
						int rowToStart = mcRowCounter;
						System.out.println("row to start the count is-"+rowToStart);
						rowNo = rowToStart;
						return rowToStart;
						// here we start with the row 
						// put these variables in the global list of variables
					
										
				}
				else
				{
					// We can still be specific with error message by having 2 if conditions 1 for Y  and other for START
					System.out.println("start point not found");
				}
				
				
			}
			
		}
		catch(Exception e)
		{
			System.out.println("Exception Found in finding the file "+ e);
			return null;
		}
	
	return rowNo;
	}

}
