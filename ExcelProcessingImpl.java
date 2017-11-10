package services;


import utils.ResourceUtil;
import java.io.File;
import java.io.FileOutputStream;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelProcessingImpl {
	
	private static final Logger log = Logger.getLogger(GeneratorInterfaceImpl.class);
	

	public static void main(String[] args) {
		System.out.println("Start:");
		ExcelProcessingImpl ReportGenerator = new ExcelProcessingImpl();
		ReportGenerator.startThread();
		System.out.println("End.");
		
	}


	public void startThread() {
		try {
			System.out.println("Start startThread() of ExcelProcessingImpl");
			processExportExcel(ResourceUtil.getCommonProperty("file.sourceDirectory"));
			System.out.println("successfully process the excel files!");
			System.out.println("Finish startThead() of ExcelProcessingImpl!");
		} catch(Exception e) {
			e.printStackTrace();
		}
		
	}
	
	
	public String processExportExcel(String sourceDirectory) {
		System.out.println("Start Process on processExportExcel..");
		String returnMsg = "";
		int[] session1 = new int[6];
		int[] session2 = new int[6];
		int[] session3 = new int[6];
		
		try {
			
			File rootDir = new File(sourceDirectory + "/UL2000Refarming");
		
			for (File excel : rootDir.listFiles()) {
				
				System.out.println("\nStart to read xlsx files...");
				
				if (excel.getName().endsWith("xlsx")) {
					System.out.println("Processing on:: " + excel.getName());
					
					//Import the excel
					Workbook wb = WorkbookFactory.create(excel);
					Sheet sheet1, sheet2, sheet3;
					XSSFSheet outputSheet1, outputSheet2, outputSheet3;
					
					XSSFWorkbook writewb = new XSSFWorkbook();
					
					//Process on the First sheet
					System.out.println("\nProcess on the first sheet");
					sheet1 = wb.getSheetAt(0);
					outputSheet1 = writewb.createSheet("Download");
					session1 = processSheet(sheet1, outputSheet1, wb, writewb);
					
					//Process on the second sheet
					System.out.println("\nProcess on the second sheet");
					sheet2 = wb.getSheetAt(1);
					outputSheet2 = writewb.createSheet("Upload");
					session2 = processSheet(sheet2, outputSheet2, wb, writewb);
					
					//Process on the third sheet
					System.out.println("\nProcess on the third sheet");
					sheet3 = wb.getSheetAt(2);
					outputSheet3 = writewb.createSheet("Ping");
					session3 = processSheet(sheet3, outputSheet3, wb, writewb);
					
					//Add the First Summary sheet	
					
					printArray(session1);
					printArray(session2);
					printArray(session3);
					
					Sheet sheet = writewb.createSheet("Summary");

					Row RowInstance;
					Cell cell = null;
					
					

					//row 1
					RowInstance = sheet.createRow(0);
					
					for (int i = 0; i<5; i++)
						cell = RowInstance.createCell(i);
					cell.setCellValue(excel.getName());
					sheet.addMergedRegion(new CellRangeAddress(0,0,0,4));
					
					//row 2
					RowInstance = sheet.createRow(1);
					//Description	Session 1	Session 2	Session 3	Session 4	Session 5
					cell = RowInstance.createCell(0);
					cell.setCellValue("Description");
					cell = RowInstance.createCell(1);
					cell.setCellValue("Session 1");
					cell = RowInstance.createCell(2);
					cell.setCellValue("Session 2");
					cell = RowInstance.createCell(3);
					cell.setCellValue("Session 3");
					cell = RowInstance.createCell(4);
					cell.setCellValue("Session 4");
					cell = RowInstance.createCell(5);
					cell.setCellValue("Session 5");
					
					
					
					//row 3
					RowInstance = sheet.createRow(2);
					cell = RowInstance.createCell(0);
					cell.setCellValue("UARFCN");
					
					for(int i = 1; i<=5; i++) {
						cell = RowInstance.createCell(i);
						String str = "AVERAGE(Download!C"+session1[i-1]+":C"+session1[i]+")";
						cell.setCellFormula(str);
					}
					
					
					//row 4
					RowInstance = sheet.createRow(3);
					cell = RowInstance.createCell(0);
					cell.setCellValue("DL Average RSCP (dBm)");
				
					for(int i = 1; i<=5; i++) {
						cell = RowInstance.createCell(i);
						cell.setCellFormula("AVERAGE(Download!E"+session1[i-1]+":E"+session1[i]+")");
					}
					
					
					//row 5
					RowInstance = sheet.createRow(4);
					cell = RowInstance.createCell(0);
					cell.setCellValue("DL Average Ec/N0 (dB)");
					
					for(int i = 1; i<=5; i++) {
						cell = RowInstance.createCell(i);
						cell.setCellFormula("AVERAGE(Download!F"+session1[i-1]+":F"+session1[i]+")");
					}
					
					
					
					//row 6
					RowInstance = sheet.createRow(5);
					cell = RowInstance.createCell(0);
					cell.setCellValue("Average DL Throughput (Mbps)");

					for(int i = 1; i<=5; i++) {
						cell = RowInstance.createCell(i);
						cell.setCellFormula("AVERAGE(Download!G"+session1[i-1]+":G"+session1[i]+")*0.000001");
					}
					
					
					//row 7
					RowInstance = sheet.createRow(6);
					cell = RowInstance.createCell(0);
					cell.setCellValue("Peak DL Throughput (Mbps)");

					for(int i = 1; i<=5; i++) {
						cell = RowInstance.createCell(i);
						cell.setCellFormula("PERCENTILE(Download!G"+session1[i-1]+":G"+session1[i]+",90%)*0.000001");
					}
					
					
					//row 8
					RowInstance = sheet.createRow(7);
					cell = RowInstance.createCell(0);
					cell.setCellValue("Average UL Throughput (Mbps)");

					for(int i = 1; i<=5; i++) {
						cell = RowInstance.createCell(i);
						cell.setCellFormula("AVERAGE(Upload!I"+session2[i-1]+":I"+session2[i]+")*0.000001");
					}
					
					
					//row 9
					RowInstance = sheet.createRow(8);
					cell = RowInstance.createCell(0);
					cell.setCellValue("Peak UL Throughput (Mbps)");

					for(int i = 1; i<=5; i++) {
						cell = RowInstance.createCell(i);
						cell.setCellFormula("PERCENTILE(Upload!I"+session2[i-1]+":I"+session2[i]+",90%)*0.000001");
					}
					
					
					//row 10
					RowInstance = sheet.createRow(9);
					cell = RowInstance.createCell(0);
					cell.setCellValue("Average Ping Round Trip Time (ms)");

					for(int i = 1; i<=5; i++) {
						cell = RowInstance.createCell(i);
						cell.setCellFormula("AVERAGE(Ping!J"+session3[i-1]+":J"+session3[i]+")");
					}
					
					///////////////////////////////////////
					//Output
					File file = new File(sourceDirectory + "/Output/" + excel.getName().split("\\.")[0] + "_Processed.xlsx");
					if (!file.exists()) {
						file.createNewFile();
					}
					
					FileOutputStream out = new FileOutputStream(file);
					writewb.write(out);
					out.close();
					writewb.close();
					wb.close();
				}
				
			}
		
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return returnMsg;
	}
	
	boolean compareTime (Cell pre, Cell current) {
		
		ImportReportInterfaceImpl importInterfaceImpl = new ImportReportInterfaceImpl();
		
		
		double currentTime = current.getNumericCellValue();
		double preTime = pre.getNumericCellValue();
		
		double diff = currentTime - preTime;
		
		if (diff > 0.0005) return true;
		else return false;
		
		
	}
	
	int[] processSheet (Sheet sheet, XSSFSheet outputSheet, Workbook originWorkbook, Workbook outputWorkbook) throws Exception{
		
		int[] sessionID = new int[6];
		sessionID[0] = 2;
	
		Row currentRow;
		
		if (sheet == null) {
			throw new Exception("EmptySheet");
		}
		
		int rowNum = sheet.getPhysicalNumberOfRows();
		System.out.println("==== No. of rows = " + rowNum);
		System.out.println("\nValue: 0" + "  SessionID: " + sessionID[0]);
		
		copySheet(outputSheet, sheet, 1, outputWorkbook, originWorkbook);
			//First row first column, "Session"
		outputSheet.getRow(0).createCell(0).setCellValue("Session");
		
		currentRow = outputSheet.getRow(0);
		Cell cell = currentRow.getCell(0);
		
		if (cell == null) sheet.getRow(0).createCell(0).setCellValue("Session");
		else cell.setCellValue("Session");
		
			//deal with first Line first
		
		currentRow = outputSheet.getRow(1);
		cell = currentRow.getCell(0);
		if (cell == null) currentRow.createCell(0).setCellValue(1);
		else cell.setCellValue(1);
		
		boolean flag = true;
		Cell compareCell;
		int value = 1;
		for (int i=2; i<rowNum; i++) {
			
			if (currentRow != null) {
			
				compareCell = currentRow.getCell(1);
				currentRow = outputSheet.getRow(i);

				//compare between current and pre cell to see if needs to change the session ID
				if (compareTime(compareCell, currentRow.getCell(1)))
					flag = false;
				
				if (!flag) {
				
					sessionID[value] = i;
					System.out.println("Value: " + value + "  SessionID: " + sessionID[value]);
					value++; flag = true;
				
				}
				
				cell = currentRow.getCell(0);
				if (cell == null) currentRow.createCell(0).setCellValue(value);
				else cell.setCellValue(value);
				
			}
			else {
				throw new Exception("rowNum error.");
			}
			
		}
		
		sessionID[5] = rowNum;
		return sessionID;	
	}
	
	void printArray(int[] array) {
		for(int i = 0; i<=5; i++)
			System.out.print(array[i] + ", ");
		System.out.println("\n");
	}
	
	void copySheet(Sheet targetSheet, Sheet originSheet, int colIndex, Workbook targetWorkbook, Workbook originWorkbook) {
		
		ImportReportInterfaceImpl importInterface = new ImportReportInterfaceImpl();
		//style clone
		
		int rowNum = originSheet.getPhysicalNumberOfRows();
		int colNum;
		Cell cellInstance, cell;
		Row rowInstance = originSheet.getRow(1);
		Row rowTarget;
		
		CellStyle style = targetWorkbook.createCellStyle();
		CellStyle originStyle = rowInstance.getCell(0).getCellStyle();
		style.cloneStyleFrom(originStyle);
		
		
		for (int i = 0; i < rowNum; i++) {
			
			rowInstance = originSheet.getRow(i);
			colNum = rowInstance.getLastCellNum();
			
			rowTarget = targetSheet.createRow(i);
			for (int j = colIndex; j <= colNum; j++) {
				cellInstance = rowTarget.createCell(j);
				if (i > 0) {
					cell = rowInstance.getCell(j-colIndex);
					if (cell != null) {
						//Numeric Style
						if (cell.getCellTypeEnum().compareTo(CellType.NUMERIC) == 0) cellInstance.setCellValue(cell.getNumericCellValue());
						//Time style
						if (j == colIndex) cellInstance.setCellStyle(style);
						//String Style
						if (cell.getCellTypeEnum().compareTo(CellType.STRING) == 0) cellInstance.setCellValue(cell.getStringCellValue());
					}
				}
				else cellInstance.setCellValue(importInterface.returnCellValue(rowInstance.getCell(j-colIndex)));
				
			}
		}
		
	}
}