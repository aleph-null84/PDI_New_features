package yuriy.tests;

import yuriy.release.*;

import java.io.File;

public class ExcelRemoveRows {

	public static void main(String[] args) {
		File file = new File("");
		String filePath = file.getAbsolutePath();
		
		System.out.println("Excel 2003");
		
		filePath += File.separator + "output" + File.separator; 
		String fileInput = filePath + "testFile.xls";
		String fileOutput = filePath + "testFile_1.xls";
		
		int i = ExcelFunctions.excelRemoveRows(fileInput, fileOutput, false, 0, 5, 4);
		System.out.println("Rows deleted: " + i);
		
		System.out.println("Excel 2007");
		
		filePath = file.getAbsolutePath();
		filePath += File.separator + "output" + File.separator; 
		fileInput = filePath + "testFile.xlsx";
		fileOutput = filePath + "testFile_1.xlsx";
		
		i = ExcelFunctions.excelRemoveRows(fileInput, fileOutput, true, 0, 5, 4);
		System.out.println("Rows deleted: " + i);
	}
}
