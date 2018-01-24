package services;

import java.io.BufferedWriter;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Scanner;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import utils.ResourceUtil;



public class ExportReportInterfaceImpl implements ExportReportInterface {
	static int linenum = 1;
	private static final Logger log = Logger.getLogger(ExportReportInterfaceImpl.class);
	
	public Scanner LOCUSExportToCSVByScanner(Scanner sc, BufferedWriter output) throws IOException {
		
		int exportRowNum = 0;
		
		
		System.out.println("Starting to set Export File-Contents in First Sheet..");
		setExportFileContentandHeaders(output);
		
		System.out.println("Getting on write rows...");
		sc.nextLine();
		while(sc.hasNext()) {
			try {
				if (exportRowNum >= Integer.parseInt(ResourceUtil.getCommonProperty("file.CSVLimits"))) break;
			} catch (NumberFormatException e) {
				e.printStackTrace();
			} catch (Exception e) {
				e.printStackTrace();
			}
			
			String line = sc.nextLine();
			String[] cells = line.split("\t");
			linenum++;
			
			if (!checkValidationForStringArray(cells)) continue;
			
			exportRowNum++;
			
			//MCC
			output.write("525");
	    
			//MNC
			output.write(",");
			output.write("03");
			
			//MRTime
			output.write(",");
			output.write(cells[0]);
			
			//lon. lat.
			output.write(",");
			output.write(cells[2]);
			
			output.write(",");
			output.write(cells[1]);
			
			//eNodeBID, CellID
			output.write(",");
			output.write(String.valueOf(geteNodeBID(cells[3])));
			
			output.write(",");
			output.write(String.valueOf(getCellID(cells[3])));
			
			//SEARFCN	SPCI	SRSCP

			output.write(",");
			output.write((cells[8]));
			 
			output.write(",");
			output.write((cells[9]));
			
			output.write(",");
			double val = Double.parseDouble(cells[10]);
			DecimalFormat df = new DecimalFormat("#.##");
			String cellString = df.format(val);
			output.write(cellString);
			
			
			output.write(",,");
			int i = 0;
			//NEARFCN	NPCI	NRSCP
			while((13+i*6)<cells.length) {
				
				output.write(",");
				output.write(cells[15 + i*6]);
				
				output.write(",");
				output.write(cells[16 + i*6]);
				
				output.write(",");
				if (cells[13 + i*6].equalsIgnoreCase(""))
					output.write("");
				else {
					
					BigDecimal bd = new BigDecimal(cells[13 + i*6]);
					bd = bd.setScale(3,RoundingMode.HALF_UP);
					cellString = bd.toPlainString();
					output.write(cellString);
					
				}
				
				output.write(",,");
				i++;
				
			}
			
			output.write(System.getProperty("line.separator"));
			
		}		
		
		return sc;
	}
	public Scanner LOCUSExportToFirstSheetByScanner(String sheetName, Scanner sc, SXSSFWorkbook writewb) {
		
		
		SXSSFSheet spreadSheet = writewb.createSheet(sheetName);
		writewb.setSheetOrder(sheetName, writewb.getSheetIndex(spreadSheet));
		
		Row exportRow = spreadSheet.createRow(0);
		Cell cell = null;
		int cellnum = 0;
		int exportRowNum = 1;

		System.out.println("Starting to set Export File-Contents in First Sheet..");
		exportRow =	setExportFileContentandHeaders(writewb, cell, exportRow);
	
		
		System.out.println("Getting on write rows...");
		sc.nextLine();
		while (sc.hasNext()) {
			
			try {
				if (exportRowNum >= Integer.parseInt(ResourceUtil.getCommonProperty("file.excelLimits"))) break;
			} catch (NumberFormatException e) {
				e.printStackTrace();
			} catch (Exception e) {
				e.printStackTrace();
			}
			
			String line = sc.nextLine();
			String[] cells = line.split("\t");
			linenum++;
			
			if (!checkValidationForStringArray(cells)) continue;
			
			exportRow = spreadSheet.createRow(exportRowNum++);
			
			//MCC
			cell = exportRow.createCell(cellnum++);
			cell.setCellValue(525);
	    
			//MNC
			cell = exportRow.createCell(cellnum++);
			cell.setCellValue("03");
			
			//MRTime
			cell = exportRow.createCell(cellnum++);
			cell.setCellValue(cells[0]);
			
			//lon. lat.
			cell = exportRow.createCell(cellnum++);
			cell.setCellValue(cells[2]);
			
			cell = exportRow.createCell(cellnum++);
			cell.setCellValue(cells[1]);
			
			//eNodeBID, CellID
			cell = exportRow.createCell(cellnum++);
			cell.setCellValue(geteNodeBID(cells[3]));
			
			cell = exportRow.createCell(cellnum++);
			cell.setCellValue(getCellID(cells[3]));
			
			//SEARFCN	SPCI	SRSCP

			cell = exportRow.createCell(cellnum++);
			cell.setCellValue((cells[8]));
			
			cell = exportRow.createCell(cellnum++);
			cell.setCellValue((cells[9]));
			
			cell = exportRow.createCell(cellnum++);
			double val = Double.parseDouble(cells[10]);
			DecimalFormat df = new DecimalFormat("###.00");
			String cellString = df.format(val);
			cell.setCellValue(cellString);
			
			
			cellnum += 2;
			int i = 0;
			//NEARFCN	NPCI	NRSCP
			while((13+i*6)<cells.length) {
				
				cell = exportRow.createCell(cellnum++);
				cell.setCellValue(cells[15 + i*6]);
				
				cell = exportRow.createCell(cellnum++);
				cell.setCellValue(cells[16 + i*6]);
				
				cell = exportRow.createCell(cellnum++);
				if (cells[13 + i*6].equalsIgnoreCase(""))
					cell.setCellValue("");
				else {
					val = Double.parseDouble(cells[13 + i*6]);
					df = new DecimalFormat("###.00");
					cellString = df.format(val);
					cell.setCellValue(cellString);
				}
				
				cellnum += 2;
				i++;
				
			}
			
			cellnum = 0;
			
		}
		
		
		return sc;
	}

	public boolean checkValidationForStringArray(String[] cells) {
		
		if (cells.length<15) return false;
		
		if (cells[1].equalsIgnoreCase("0") || cells[2].equalsIgnoreCase("0"))
			return false;
		
		for(int i = 1; i <= 10; i++) {
			
			if (cells[i].equalsIgnoreCase("")) return false;
		}
		
		return true;
	}

	public String LOCUSExportToFirstSheetBySheet(String sheetName, SXSSFWorkbook writewb, Sheet importSheet ) {
		
		String returnMsg = "";
		ImportReportInterface importReportInterfaceImpl = new ImportReportInterfaceImpl();
		
	
		SXSSFSheet spreadSheet = writewb.createSheet(sheetName);
		writewb.setSheetOrder(sheetName, writewb.getSheetIndex(spreadSheet));
		
		Row exportRow = spreadSheet.createRow(0);
		Cell cell = null;
		
		System.out.println("Starting to set Export File-Contents in First Sheet..");
		exportRow =	setExportFileContentandHeaders(writewb, cell, exportRow);
		
		int rownum = 1;
		int exportRowNum = 1;
		
		int rowCount = importSheet.getLastRowNum();
		int cellnum = 0;
		
		while (rownum <= rowCount) {
//			System.out.println("Written row ::" + rownum);
//			
			Row importRow = importSheet.getRow(rownum);
			
			if (!checkValidationForExcelRow(importRow,rownum)) {
				 rownum++;
			     continue;
			}
			exportRow = spreadSheet.createRow(exportRowNum++);
			
			//MCC
			cell = exportRow.createCell(cellnum++);
			cell.setCellValue(525);
	    
			//MNC
			cell = exportRow.createCell(cellnum++);
			cell.setCellValue("03");
			
			//MRTime
			cell = exportRow.createCell(cellnum++);
			CellStyle style = importRow.getCell(0).getCellStyle();
			CellStyle exportStyle = writewb.createCellStyle();
			exportStyle.cloneStyleFrom(style);
			cell.setCellStyle(exportStyle);
			cell.setCellValue(importRow.getCell(0).getNumericCellValue());
			
			//lon. lat.
			cell = exportRow.createCell(cellnum++);
			cell.setCellValue(importRow.getCell(2).getNumericCellValue());
			
			cell = exportRow.createCell(cellnum++);
			cell.setCellValue(importRow.getCell(1).getNumericCellValue());
			
			//eNodeBID, CellID
			cell = exportRow.createCell(cellnum++);
			cell.setCellValue(geteNodeBID(importReportInterfaceImpl.returnCellValue(importRow.getCell(3))));
			
			cell = exportRow.createCell(cellnum++);
			cell.setCellValue(getCellID(importReportInterfaceImpl.returnCellValue(importRow.getCell(3))));
			
			//SEARFCN	SPCI	SRSCP

			cell = exportRow.createCell(cellnum++);
			cell.setCellValue((importReportInterfaceImpl.returnCellValue(importRow.getCell(8))));
			
			cell = exportRow.createCell(cellnum++);
			cell.setCellValue(importReportInterfaceImpl.returnCellValue(importRow.getCell(9)));
			
			cell = exportRow.createCell(cellnum++);
			cell.setCellValue(importReportInterfaceImpl.returnCellValue(importRow.getCell(10)));
			
			cellnum += 2;
			int i = 0;
			//NEARFCN	NPCI	NRSCP
			while(importRow.getCell(13 + i*6) != null) {
				
				cell = exportRow.createCell(cellnum++);
				cell.setCellValue(importReportInterfaceImpl.returnCellValue(importRow.getCell(13 + i*6)));
				
				cell = exportRow.createCell(cellnum++);
				cell.setCellValue(importReportInterfaceImpl.returnCellValue(importRow.getCell(15 + i*6)));
				
				cell = exportRow.createCell(cellnum++);
				cell.setCellValue(importReportInterfaceImpl.returnCellValue(importRow.getCell(16 + i*6)));
				
				cellnum += 2;
				i++;
				if (importRow.getCell(13 + i*6) == null)	break;
				
			}
	        rownum++;
	        cellnum = 0;
		}
		
		
		
		returnMsg = "Successfully Exoport the Excel!";
		return returnMsg;
	}

	public boolean checkValidationForExcelRow(Row importRow, int rowNum) {
		//lon. lat. cannot be 0
		Cell cell = importRow.getCell(1,MissingCellPolicy.RETURN_BLANK_AS_NULL);
		if (cell.getNumericCellValue() == 0 || cell.toString().compareTo("") == 0) {
			System.out.println("Dispose Row ::" + (rowNum+1));
			return false;
		}
		cell = importRow.getCell(2,MissingCellPolicy.RETURN_BLANK_AS_NULL);
		if (cell.getNumericCellValue() == 0 || cell.toString().compareTo("") == 0) {
			System.out.println("Dispose Row ::" + (rowNum+1));
			return false;
		}
		
		//the rest cannot be empty
		for (int i = 3; i <= 10; i++) {
			cell = importRow.getCell(i,MissingCellPolicy.CREATE_NULL_AS_BLANK);
			if (cell.getCellTypeEnum() == CellType.BLANK) {
				System.out.println("Dispose Row ::" + (rowNum+1));
				return false;
			}
		}
		return true;
	}

	public int geteNodeBID(String input) {
		int dec = Integer.parseInt(input);
		String hex = Integer.toHexString(dec);
		int CellIDIndex = hex.length() - 2;
		String eNodeBID = hex.substring(0, CellIDIndex);
		
		return Integer.parseInt(eNodeBID, 16);
		
	}
	
	public int getCellID(String input) {
		int dec = Integer.parseInt(input);
		int CellIDIndex = 0;
		String hex = Integer.toHexString(dec);
		CellIDIndex = hex.length() - 2;
		String CellID = hex.substring(CellIDIndex);  
		
		return Integer.parseInt(CellID, 16);
		
	}

	public void setExportFileContentandHeaders(BufferedWriter output) throws IOException {
		
		int i = 0;

		output.write("MCC");
		

		output.write(",");
		output.write("MNC");

		//MRTime
		output.write(",");
		output.write("MRTime");
		
		//Longitude & latitude
		output.write(",");
		output.write("Longitude");
		
		output.write(",");
		output.write("Latitude");
		
		//eNodeBID & CellID
		output.write(",");
		output.write("eNodeBID");
		
		output.write(",");
		output.write("CellID");
		
		//SEARFCN
		output.write(",");
		output.write("SEARFCN");
		
		output.write(",");
		output.write("SPCI");
		
		output.write(",");
		output.write("SRSCP");

		///////////////////////////////////////////////////
		output.write(",");
		output.write("eNodeBID1");

		output.write(",");
		output.write("CellID1");

		output.write(",");
		output.write("NEARFCN1");

		output.write(",");
		output.write("NPCI1");

		output.write(",");
		output.write("NRSCP1");

		output.write(",");
		output.write("eNodeBID2");
		
		output.write(",");
		output.write("CellID2");

		output.write(",");
		output.write("NEARFCN2");
	
		output.write(",");
		output.write("NPCI2");

		output.write(",");
		output.write("NRSCP2");

		output.write(",");
		output.write("eNodeBID3");

		output.write(",");
		output.write("CellID3");
		
		output.write(",");
		output.write("NEARFCN3");
		
		output.write(",");
		output.write("NPCI3");
		
		output.write(",");
		output.write("NRSCP3");
		
		//eNodeBID4	CellID4	NEARFCN4	NPCI4	NRSCP4	eNodeBID5	CellID5	NEARFCN5	

		output.write(",");
		output.write("eNodeBID4");

		output.write(",");
		output.write("CellID4");

		output.write(",");
		output.write("NEARFCN4");

		output.write(",");
		output.write("NPCI4");

		output.write(",");
		output.write("NRSCP4");

		output.write(",");
		output.write("eNodeBID5");

		output.write(",");
		output.write("CellID5");

		output.write(",");
		output.write("NEARFCN5");

		output.write(",");
		output.write("NPCI5");

		////NPCI5	NRSCP5	eNodeBID6	CellID6	NEARFCN6	NPCI6	NRSCP6	eNodeBID7	CellID7	NEARFCN7	NPCI7	
		output.write(",");
		output.write("NRSCP5");

		output.write(",");
		output.write("eNodeBID6");

		output.write(",");
		output.write("CellID6");
		
		output.write(",");
		output.write("NEARFCN6");

		output.write(",");
		output.write("NPCI6");

		output.write(",");
		output.write("NRSCP6");

		output.write(",");
		output.write("eNodeBID7");

		output.write(",");
		output.write("CellID7");

		output.write(",");
		output.write("NEARFCN7");

		output.write(",");
		output.write("NPCI7");
		
		///NRSCP7	eNodeBID8	CellID8	NEARFCN8	NPCI8	NRSCP8

		output.write(",");
		output.write("NRSCP7");

		output.write(",");
		output.write("eNodeBID8");

		output.write(",");
		output.write("CellID8");

		output.write(",");
		output.write("NEARFCN8");

		output.write(",");
		output.write("NPCI8");
		
		output.write(",");
		output.write("NRSCP8");
		
		output.write(System.getProperty("line.separator"));
		System.out.println("=====Finish setting header names\n");
		return;
		
	}
	public Row setExportFileContentandHeaders(Workbook wb, Cell cell, Row row) {
		System.out.println("=====Inside setExportFileContentandHeaders Setting the Header Names..");
		
//		style.setBorderTop((short) 2);
//		style.setBorderRight(HSSFCellStyle.BORDER_THIN);
//		style.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
//		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//		style.setAlignment(CellStyle.ALIGN_CENTER);
//		Font defaultFont = workbook.createFont();
//		defaultFont.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
//		style.setFont(defaultFont);
		
		int i = 0;
		
		cell = row.createCell(i++);
		cell.setCellValue("MCC");
		

		cell = row.createCell(i++);
		cell.setCellValue("MNC");

		//MRTime
		cell = row.createCell(i++);
		cell.setCellValue("MRTime");
		
		//Longitude & latitude
		cell = row.createCell(i++);
		cell.setCellValue("Longitude");
		
		cell = row.createCell(i++);
		cell.setCellValue("Latitude");
		
		//eNodeBID & CellID
		cell = row.createCell(i++);
		cell.setCellValue("eNodeBID");
		
		cell = row.createCell(i++);
		cell.setCellValue("CellID");
		
		//SEARFCN
		cell = row.createCell(i++);
		cell.setCellValue("SEARFCN");
		
		cell = row.createCell(i++);
		cell.setCellValue("SPCI");
		
		cell = row.createCell(i++);
		cell.setCellValue("SRSCP");

		///////////////////////////////////////////////////
		cell = row.createCell(i++);
		cell.setCellValue("eNodeBID1");

		cell = row.createCell(i++);
		cell.setCellValue("CellID1");

		cell = row.createCell(i++);
		cell.setCellValue("NEARFCN1");

		cell = row.createCell(i++);
		cell.setCellValue("NPCI1");

		cell = row.createCell(i++);
		cell.setCellValue("NRSCP1");

		cell = row.createCell(i++);
		cell.setCellValue("eNodeBID2");
		
		cell = row.createCell(i++);
		cell.setCellValue("CellID2");

		cell = row.createCell(i++);
		cell.setCellValue("NEARFCN2");
	
		cell = row.createCell(i++);
		cell.setCellValue("NPCI2");

		cell = row.createCell(i++);
		cell.setCellValue("NRSCP2");

		cell = row.createCell(i++);
		cell.setCellValue("eNodeBID3");

		cell = row.createCell(i++);
		cell.setCellValue("CellID3");
		
		cell = row.createCell(i++);
		cell.setCellValue("NEARFCN3");
		
		cell = row.createCell(i++);
		cell.setCellValue("NPCI3");
		
		cell = row.createCell(i++);
		cell.setCellValue("NRSCP3");
		
		//eNodeBID4	CellID4	NEARFCN4	NPCI4	NRSCP4	eNodeBID5	CellID5	NEARFCN5	

		cell = row.createCell(i++);
		cell.setCellValue("eNodeBID4");

		cell = row.createCell(i++);
		cell.setCellValue("CellID4");

		cell = row.createCell(i++);
		cell.setCellValue("NEARFCN4");

		cell = row.createCell(i++);
		cell.setCellValue("NPCI4");

		cell = row.createCell(i++);
		cell.setCellValue("NRSCP4");

		cell = row.createCell(i++);
		cell.setCellValue("eNodeBID5");

		cell = row.createCell(i++);
		cell.setCellValue("CellID5");

		cell = row.createCell(i++);
		cell.setCellValue("NEARFCN5");

		cell = row.createCell(i++);
		cell.setCellValue("NPCI5");

		////NPCI5	NRSCP5	eNodeBID6	CellID6	NEARFCN6	NPCI6	NRSCP6	eNodeBID7	CellID7	NEARFCN7	NPCI7	
		cell = row.createCell(i++);
		cell.setCellValue("NRSCP5");

		cell = row.createCell(i++);
		cell.setCellValue("eNodeBID6");

		cell = row.createCell(i++);
		cell.setCellValue("CellID6");

		cell = row.createCell(i++);
		cell.setCellValue("NPCI6");

		cell = row.createCell(i++);
		cell.setCellValue("NRSCP6");

		cell = row.createCell(i++);
		cell.setCellValue("eNodeBID7");

		cell = row.createCell(i++);
		cell.setCellValue("CellID7");

		cell = row.createCell(i++);
		cell.setCellValue("NEARFCN7");

		cell = row.createCell(i++);
		cell.setCellValue("NPCI7");
		
		///NRSCP7	eNodeBID8	CellID8	NEARFCN8	NPCI8	NRSCP8

		cell = row.createCell(i++);
		cell.setCellValue("NRSCP7");

		cell = row.createCell(i++);
		cell.setCellValue("eNodeBID8");

		cell = row.createCell(i++);
		cell.setCellValue("CellID8");

		cell = row.createCell(i++);
		cell.setCellValue("NEARFCN8");

		cell = row.createCell(i++);
		cell.setCellValue("NPCI8");
		
		cell = row.createCell(i++);
		cell.setCellValue("NRSCP8");
		
		System.out.println("=====Finish setting header names\n");
		return row;
	}
	
}