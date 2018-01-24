package services;

import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public interface ExportReportInterface {

	public abstract Scanner LOCUSExportToFirstSheetByScanner(String sheetName,
			Scanner sc, SXSSFWorkbook writewb);

	public abstract boolean checkValidationForStringArray(String[] cells);

	public abstract String LOCUSExportToFirstSheetBySheet(String sheetName,
			SXSSFWorkbook writewb, Sheet importSheet);

	public abstract boolean checkValidationForExcelRow(Row importRow, int rowNum);

	public abstract int geteNodeBID(String input);

	public abstract int getCellID(String input);

	public abstract Row setExportFileContentandHeaders(Workbook wb, Cell cell,
			Row row);

}