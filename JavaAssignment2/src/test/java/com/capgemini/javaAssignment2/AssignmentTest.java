package com.capgemini.javaAssignment2;

import java.io.File;
import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;

public class AssignmentTest {
	@Test
	public void differentNumberOfSheets() throws EncryptedDocumentException, IOException
	{
	String userDir = System.getProperty("user.dir");
	Workbook wb1 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\DataFiles\\ExcelFilesWithThreeSheets.xlsx"));
	Workbook wb2 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\DataFiles\\ExcelFilesWithTwoSheets-Data-Actual.xlsx"));
	JavaAssignment2 Assignment2 = new JavaAssignment2();
	Assignment2.verifyIfExcelFilesHaveSameNumberAndNameOfSheets(wb1, wb2);

	}

	@Test
	public void sameNumberOfSheets() throws EncryptedDocumentException, IOException
	{
	String userDir = System.getProperty("user.dir");
	Workbook wb1 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\DataFiles\\ExcelFilesWithTwoSheets-Data-Actual.xlsx"));
	Workbook wb2 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\DataFiles\\ExcelFilesWithTwoSheets-Data-Match.xlsx"));
	JavaAssignment2 Assignment2 = new JavaAssignment2();
	Assignment2.verifyIfExcelFilesHaveSameNumberAndNameOfSheets(wb1, wb2);
	}


	@Test
	public void differentNumberOfRows() throws EncryptedDocumentException, IOException
	{
	String userDir = System.getProperty("user.dir");
	Workbook wb1 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\DataFiles\\ExcelFilesWithTwoSheets-Data-Actual.xlsx"));
	Workbook wb2 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\DataFiles\\ExcelFilesWithTwoSheets-Extra-Row.xlsx"));
	JavaAssignment2 Assignment2 = new JavaAssignment2();
	Assignment2.verifySheetsInExcelFilesHaveSameRowsAndColumns(wb1, wb2);
	}

	@Test
	public void differentNumberOfColumns() throws EncryptedDocumentException, IOException
	{
	String userDir = System.getProperty("user.dir");
	Workbook wb1 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\DataFiles\\ExcelFilesWithTwoSheets-Data-Actual.xlsx"));
	Workbook wb2 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\DataFiles\\ExcelFilesWithTwoSheets-Extra-Columns.xlsx"));
	JavaAssignment2 Assignment2 = new JavaAssignment2();
	Assignment2.verifySheetsInExcelFilesHaveSameRowsAndColumns(wb1, wb2);
	}

	@Test
	public void sameContentOfExcelFiles() throws EncryptedDocumentException, IOException {
	String userDir = System.getProperty("user.dir");
	Workbook wb1 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\DataFiles\\ExcelFilesWithTwoSheets-Data-Actual.xlsx"));
	Workbook wb2 = WorkbookFactory.create(new File(userDir+"\\src\\test\\resources\\DataFiles\\ExcelFilesWithTwoSheets-Data-Match.xlsx"));
	JavaAssignment2 Assignment2 = new JavaAssignment2();
	Assignment2.verifyIfExcelFilesHaveSameNumberAndNameOfSheets(wb1, wb2);
	Assignment2.verifySheetsInExcelFilesHaveSameRowsAndColumns(wb1, wb2);
	Assignment2.verifyDataInExcelBookAllSheets(wb1, wb2);

	}
}
