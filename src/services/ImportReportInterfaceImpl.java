package services;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Scanner;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import utils.Constants;
import utils.ResourceUtil;


public class ImportReportInterfaceImpl implements ImportReportInterface {
	private static final Logger log = Logger.getLogger(ExportReportInterfaceImpl.class);
	
	
	@Override
	public String returnCellValue(Cell cell) {
		String value = "";
		if (cell != null) {

			if (cell.getCellTypeEnum() == CellType.NUMERIC) {
				
				value = (cell != null && cell.toString().length() > 0 ? String
						.valueOf((int) Math.round(cell.getNumericCellValue()))
						: "");
				
				
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					java.util.Date utDate = cell.getDateCellValue();
					Date sqlDate = new Date(utDate.getTime());
					value = sqlDate.toString();// cell.getDateCellValue().toString();
				}
			} else if (cell.getCellTypeEnum() == CellType.STRING) {
				value = (cell != null && cell.toString().length() > 0 ? cell
						.toString() : "");
			} else if (cell.getCellTypeEnum() == CellType.FORMULA) {
				CellType fType = cell.getCachedFormulaResultTypeEnum();
				if (fType == CellType.NUMERIC) {
					
					//if got dot then represents a double type
					if(String.valueOf(cell.getNumericCellValue()).contains(".")) {
						value =  (cell != null && cell.toString().length() > 0 ? String
								.valueOf(cell.getNumericCellValue()) : "");
					}
					else {
						value = (cell != null && cell.toString().length() > 0 ? String
							.valueOf((int) Math.round(cell.getNumericCellValue()))
							: "");
					}
				} else {
//					System.out.println("fType=" + fType);
					value = (cell != null && cell.toString().length() > 0 ? cell
							.getStringCellValue() : "");
				}
			} else {
				value = (cell != null && cell.toString().length() > 0 ? cell
						.toString() : "");
			}
		} else {
			value = "";
		}
		return value;
	}


	
	@Override
	public int getTotalRowCount(int tableStartIndex, int xlsheetLastRowIndex,
			Sheet ws) {
		int rowCount = 0;
		for (int i = tableStartIndex; i<xlsheetLastRowIndex;i++) {
			Row row = ws.getRow(i);
			if (row != null) {
				Cell cell = row.getCell(0);
				if (cell == null) return rowCount;
				else if (cell.toString().compareTo("") == 0) {
					return rowCount;
				}
				rowCount++;
			}
			else {
				return rowCount;
			}
		}
		return rowCount;
	}

	@Override
	public String processImportTxt(String sourceDirectory) {
		
		ExportReportInterfaceImpl exportReportInterfaceImpl = null;
		FileInputStream inputStream = null;
		FileOutputStream out = null;
		Scanner sc = null;
		File outputFile = null;;
		XSSFWorkbook workbookTmp = null;
		SXSSFWorkbook writewb = null;

		try {
			File rootDir = new File(sourceDirectory + ResourceUtil.getCommonProperty("file.sourceFiles"));
			
			for (File txt : rootDir.listFiles()) {
				

				if (txt.getAbsolutePath().endsWith(".txt")) {
					exportReportInterfaceImpl= new ExportReportInterfaceImpl();
					
				    try {
				    	System.out.println("\n=======================================================================");
				    	System.out.println("Start to process " + txt + "\n");
						inputStream = new FileInputStream(txt);
						BufferedWriter bop = null;
						sc = new Scanner(inputStream, "UTF-8");
						int i = 1;
						
						while (sc.hasNext()) {
							//create OutputFile
//							file = new File(sourceDirectory + "/Output/" + txt.getName() + "_Processed_" + (i++) + ".xlsx");
							
							Calendar cal = Calendar.getInstance();
							SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
							outputFile = new File(sourceDirectory + "/Output/" + "LTE_huawei_" + sdf.format(cal.getTime()) + ".csv");
							System.out.println("New File creadted: " + outputFile + "\n");
							
							if (!outputFile.exists()) {
								outputFile.createNewFile();
							}
							
//							workbookTmp = new XSSFWorkbook();
//							writewb = new SXSSFWorkbook(workbookTmp,5000,true,true);
							
							bop = new BufferedWriter(new FileWriter(outputFile));
							sc = exportReportInterfaceImpl.LOCUSExportToCSVByScanner(sc, bop);
							
							if (sc.ioException() != null) {
								throw sc.ioException();
							}
							
							System.out.println("Finish writing to " + outputFile + "\n");
							
							bop.flush();
							bop.close();
							
//							out = new FileOutputStream(outputFile);
//							writewb.write(out);
//							out.flush();
//							out.close();
//							writewb.dispose();

						}
						
					} catch (Exception e) {
						e.printStackTrace();
						
					} finally {
						
					    if (inputStream != null) {
					        inputStream.close();
					    }
					    if (sc != null) {
					        sc.close();
					    }
					}
				    ExportReportInterfaceImpl.linenum = 1;
				    System.out.println("Finish processing::" + txt + "\n");
				}
			}
		    
		} catch (Exception e) {
			
			e.printStackTrace();
		
		} 

		
		
		
		return "Successfully process All txt Files..";
	}
	
	@Override
	public String processImportExcelFirstSheet(String sourceDirectory) {
		
		ExportReportInterfaceImpl exportInterfaceImpl = new ExportReportInterfaceImpl();
		System.out.println(String.valueOf(new Date())); 
		
		try {
			System.out.println("Start process on processImportExcelFirstSheet()!");
			
			File rootDir = new File(sourceDirectory + ResourceUtil.getCommonProperty("file.sourceFiles"));
			
			for (File excel : rootDir.listFiles()) {
				
				if (excel.getAbsolutePath().endsWith("xlsx")) {
					
					System.out.println("Processing on :::" + excel.getName());
					
					Workbook wb = WorkbookFactory.create(excel);
					System.out.println("Number of sheets::" + wb.getNumberOfSheets());
					
					String sheetName = wb.getSheetName(0);
					System.out.println(sheetName);
					Sheet ws = wb.getSheet(sheetName);
					
					System.out.println("Finish getting sheet");
					int rowNum = 0;
					int colNum = 0;
					
					rowNum = ws.getPhysicalNumberOfRows();
					System.out.println("No. of rows = " + rowNum);
					
					colNum = ws.getRow(0).getLastCellNum();
					System.out.println("No. of columns = " + colNum);
					
					wb.close();

					File file = new File(sourceDirectory + "/Output/" + excel.getName() + "_Processed.xlsx");
					if (!file.exists()) {
						file.createNewFile();
					}
					XSSFWorkbook workbook = new XSSFWorkbook();
					SXSSFWorkbook writewb = new SXSSFWorkbook(workbook,5000,true,true);
					
					exportInterfaceImpl.LOCUSExportToFirstSheetBySheet(ResourceUtil.getCommonProperty("file.firstSheet.name"), writewb, ws); 
					
					FileOutputStream out = new FileOutputStream(file);
					writewb.write(out);
					out.close();
					writewb.dispose();
				}
				
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return "Have successfully processed all the excel files in the location";
	}
	
//	public void readData(SXSSFSheet importSheet, ArrayList<NemoRecord>) {
//		int rowNum = 0;
//		int rowCount = importSheet.getLastRowNum();
//		Row currentRow;
//		NemoRecordRow[] records = new NemoRecordRow[rowCount];
//		while (rowNum++ <= rowCount) {
//			
//			currentRow = importSheet.getRow(rowNum);
//			
//			records[rowNum].setTime(currentRow.getCell(0).getStringCellValue());
//			records[]
//		
//		
//		}
//		
//	}

}