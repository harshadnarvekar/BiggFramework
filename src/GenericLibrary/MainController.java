package GenericLibrary;

//import java.io.*;
import java.io.FileInputStream;
import java.io.OutputStream;
import java.io.InputStream;
import java.io.IOException;
import java.io.File;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MainController {
	
	/**
	 * Desc - This function is the main controller where the action starts
	 * @param args
	 */
	public static void main(String[] args) throws Exception
	{
		// declaring the global variables and constants in a file 
		// configuration vriables like browser , paths
	    // looping the maincontroller check for start and Yes
		try
		{
			File mcFile = new File("D:\\MainController.xlsx"); // global and configuration file
			XSSFSheet mcSheetHandl= null;
			int mcRowCount,mcRowCounter,mcRowToStart ;
			mcSheetHandl = ExcelUtility.GetSheetHandle(mcFile, "MainControlSheet");
			// From which row to start in the MC
			mcRowToStart = ExcelUtility.getStartingExcecutionRowOfMC(mcSheetHandl);
			// total count of rows in 
			mcRowCount = mcSheetHandl.getPhysicalNumberOfRows();
			// looping now each row and executing the test cases
			for(mcRowCounter=mcRowToStart;mcRowCounter<mcRowCount;mcRowCounter++)
			{
				XSSFRow eachRow = null;
				//XSSFCell cellValueYPos, cellValueStartPos = null;
				eachRow = mcSheetHandl.getRow(mcRowCounter);
				String testCaseId, testCasePath, pauseKeyword;
				testCaseId = eachRow.getCell(3).toString();
				testCasePath = eachRow.getCell(4).toString();
				pauseKeyword = eachRow.getCell(7).toString();
				
				System.out.println("tc-"+testCaseId +"-test path -"+testCasePath+"-pause-"+pauseKeyword);
				
			}
			
		}
		catch(Exception e)
		{
			System.out.println("Exception Found in finding the file "+ e);
		}
	}

}
