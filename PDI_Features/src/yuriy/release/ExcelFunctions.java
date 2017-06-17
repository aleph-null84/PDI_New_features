package yuriy.release;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;

public abstract class ExcelFunctions {
	
	/**
	 * Copy workbook inputFile to outputFile with removed rowCount rows starting from row rowIndex.
	 * 
	 * @param inputFile    the system-dependent input file name
	 * @param outputFile   the system-dependent output file name
	 * @param excel2007    Excel format indicator; true to Excel 2007 format, false to Excel 2003 format
	 * @param sheetNumber  the sheet number (0-based)
	 * @param rowIndex     the row index (0-based) to start removing
	 * @param rowCount     the number of rows to remove
	 * @return             count removed rows if positive or 0, error otherwise; 
	 *                     (-1: IOError with inputFile;
	 *                      -2: outputFile writing error;
	 *                      -3: IOError with outputFile)
	 */
	public static int excelRemoveRows(String inputFile,
			                          String outputFile,
			                          boolean excel2007,
			                          int sheetNumber,
			                          int rowIndex, 
			                          int rowCount) {
		// Read Excel
		Workbook wb = null;
		try {
			if (excel2007){
				wb = new XSSFWorkbook(new FileInputStream(inputFile));
			} else {
				wb = new HSSFWorkbook(new FileInputStream(inputFile));
			}
		} catch (IOException e) {
			return -1; // IOError with inputFile
		}
		
		// Delete rows
		Sheet sheet = wb.getSheetAt(sheetNumber);
		int i = excel03_07RemoveRows(sheet, rowIndex, rowCount);
		
		// Write Excel
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(outputFile);
	        try {
				wb.write(out);
			} catch (IOException e) {
				return -2; // outputFile writing error
			}finally {
				out.close();
			}
		} catch (IOException e) {
			return -3; // IOError with outputFile
		} 
		return i;
	}	
	
	private static int excel03RemoveRow(HSSFSheet sheet, int rowIndex) {
		try {
			int lastRowNum = sheet.getLastRowNum();
			if(rowIndex >=0 && rowIndex < lastRowNum){
				sheet.shiftRows(rowIndex+1,lastRowNum, -1);
			}	
			if(rowIndex == lastRowNum){
				HSSFRow removingRow = sheet.getRow(rowIndex);
				if(removingRow!=null){
					sheet.removeRow(removingRow);
				}
			}
			return 1;
		} catch (Exception e) {
			return 0;
		}
	}
	
	private static int excel07RemoveRow(XSSFSheet sheet, int rowIndex) {
		try {
			int lastRowNum = sheet.getLastRowNum();
			
			if(rowIndex >=0 && rowIndex < lastRowNum){
				XSSFRow removingRow = sheet.getRow(rowIndex);
				if(removingRow != null){
					sheet.removeRow(removingRow);
					sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
				}
			}	
			if(rowIndex == lastRowNum){
				XSSFRow removingRow = sheet.getRow(rowIndex);
				if(removingRow != null){
					sheet.removeRow(removingRow);
				}
			}
			return 1;
		} catch (Exception e) {
			return 0;
		}
	}
	
	private static int excel03_07RemoveRows(Sheet sheet, int rowIndex, int rowCount) {
		if(sheet instanceof XSSFSheet){
			return excel07RemoveRows((XSSFSheet)sheet, rowIndex, rowCount);
		}else{
			return excel03RemoveRows((HSSFSheet)sheet, rowIndex, rowCount);
		}
	}
	
	private static int excel07RemoveRows(XSSFSheet sheet, int rowIndex, int rowCount) {
		if ((rowIndex < 0) || (rowCount <= 0)) return 0;
		int deletedRovs = 0;
		
        for(int i = rowCount; i > 0; i--) {
        	deletedRovs += excel07RemoveRow(sheet, rowIndex);
        }
	    
	    return deletedRovs;
	}
	
	private static int excel03RemoveRows(HSSFSheet sheet, int rowIndex, int rowCount) {
		if ((rowIndex < 0) || (rowCount <= 0)) return 0;
		int deletedRovs = 0;
		
        for(int i = rowCount; i > 0; i--) {
        	deletedRovs += excel03RemoveRow(sheet, rowIndex);
        }
	    
	    return deletedRovs;
	}
}