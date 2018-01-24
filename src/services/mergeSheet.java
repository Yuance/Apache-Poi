package services;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.log4j.Logger;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import utils.ResourceUtil;

public class mergeSheet {
	
	private static final Logger log = Logger.getLogger(GeneratorInterfaceImpl.class);
	

	public static void main(String[] args) {
		System.out.println("Start:");
		mergeSheet MergeSheet = new mergeSheet();
		MergeSheet.startThread();
		System.out.println("End.");
		
	}


	public void startThread() {
		try {
			System.out.println("Start startThread() of mergeSheet");
			processMergeExcel(ResourceUtil.getCommonProperty("file.sourceDirectory"));
			System.out.println("successfully process the excel files!");
			System.out.println("Finish startThead() of ExcelProcessingImpl!");
		} catch(Exception e) {
			e.printStackTrace();
		}
		
	}

	public void processMergeExcel(String sourceDirectory) {
		
		File rootDir = new File(sourceDirectory + "/MergeExcels");
		
		boolean first = true;
		XSSFWorkbook writewb = new XSSFWorkbook();
		Sheet writeSheet = writewb.createSheet();
		Workbook wb;
		Sheet sheet;
		try {
			for (File excel : rootDir.listFiles()) {
				
				System.out.println("\n Start to read xlsx files...");
				
				if (excel.getName().endsWith("xlsx")) {
					System.out.println("Processing on:: " + excel.getName());
					//Import the excel
					
					wb = WorkbookFactory.create(excel);
					//Get the only sheet
					sheet = wb.getSheetAt(0);
					
					if (first) {
						setTitle(writeSheet, sheet);
						first = false;
					}
					
//					appendSheet(writeSheet, sheet, writewb, wb);
					
					wb.close();
				}
			}
			//Output
			
			File file = new File(sourceDirectory + "/Output/" +  "Combined.xlsx");
			if (!file.exists()) {
				file.createNewFile();
			}
			
			FileOutputStream out = new FileOutputStream(file);
			writewb.write(out);
			out.close();
			writewb.close();
		} catch (Exception e){
			e.printStackTrace();
		}
		
	}
	
	void setTitle(Sheet writeSheet, Sheet sheet) {
		
		Map<Integer, CellStyle> styleMap = new HashMap<Integer, CellStyle>();
		
		//Use the header of the first Sheet as the output header
		Row writeRow = writeSheet.createRow(0);
		Cell writeCell;
		Cell cell;
		for (int i = 0; i < sheet.getRow(1).getLastCellNum(); i++) {
			
			writeCell = writeRow.createCell(i);
			cell = sheet.getRow(1).getCell(i);
			copyCell(cell, writeCell, styleMap);
		}
	}
	
	void appendSheet(Sheet writeSheet, Sheet sheet, Workbook writewb, Workbook wb) {
		
		int writerowNum = writeSheet.getPhysicalNumberOfRows();
		int rowNum = sheet.getPhysicalNumberOfRows();
		int colNum;
		
		System.out.println(writerowNum + ", " + rowNum);
		
		Row writeRow;
		Row row;
		Cell cell;
		Cell writeCell;
//		CellStyle writeStyle;
//		CellStyle style;
		Map<Integer, CellStyle> styleMap = new HashMap<Integer,CellStyle>();
		
		for (int i = 2; i<rowNum; i++) {
			
			writeRow = writeSheet.createRow(writerowNum++);
			row = sheet.getRow(i);
			colNum = row.getLastCellNum();
			for (int j = 0; j<colNum; j++) {
				cell = row.getCell(j);
				writeCell = writeRow.createCell(j);
				
				copyCell(cell, writeCell, styleMap);
//				if (cell != null) {
//					// if numeric
//					if (cell.getCellTypeEnum().compareTo(CellType.NUMERIC) == 0)
//						writeCell.setCellValue(cell.getNumericCellValue());
//					// if String
//					else if (cell.getCellTypeEnum().compareTo(CellType.STRING) == 0)
//						writeCell.setCellValue(cell.getStringCellValue());
//					
//					//Style
//					style = cell.getCellStyle();
//					for (int k = 0; k<writewb.getNumCellStyles(); k++) {
//						
//					}
//					style = writewb.createCellStyle();
//					style.cloneStyleFrom();
//					writeCell.setCellStyle(style);
				}
			}
			
	}
	
	public static void copyCell(Cell oldCell, Cell newCell, Map<Integer, CellStyle> styleMap) {
        if(styleMap != null) {
            if(oldCell.getSheet().getWorkbook().equals(newCell.getSheet().getWorkbook())){
                newCell.setCellStyle(oldCell.getCellStyle());
            } else{
                int stHashCode = oldCell.getCellStyle().hashCode();
                CellStyle newCellStyle = styleMap.get(stHashCode);
                if(newCellStyle == null){
                    newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();
                    newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
                    styleMap.put(stHashCode, newCellStyle);
                }
                newCell.setCellStyle(newCellStyle);
            }
        }
      
        switch(oldCell.getCellTypeEnum()) {
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case BLANK:
                newCell.setCellType(CellType.BLANK);
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case ERROR:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            default:
                break;
        }
		
	}
}