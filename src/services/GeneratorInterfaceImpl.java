package services;



import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;

import utils.ResourceUtil;

public class GeneratorInterfaceImpl implements GeneratorInterface {
	
	private static final Logger log = Logger.getLogger(GeneratorInterfaceImpl.class);
	

	@Override
	public String processExportExcel(String sourceDirectory) {
		System.out.println("Start Process to Export Excel..");
		String returnMsg = "";
		String processMsg = "";
		try {
			
			ImportReportInterfaceImpl importInterfaceImpl = new ImportReportInterfaceImpl();
			
			processMsg = importInterfaceImpl.processImportTxt(sourceDirectory);
			
			returnMsg = "Successfully finish process on processExportExcel()!";
		
		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println(processMsg);
		return returnMsg;
	}


	@Override
	public void startThread() {
		try {
			System.out.println("Start startThread() of GeneratorInterfaceImpl!");
			processExportExcel(ResourceUtil.getCommonProperty("file.sourceDirectory"));
			System.out.println("successfully process the Txt files!");
			System.out.println("Finish startThead() of GeneratorInterfaceImpl!");
		} catch(Exception e) {
			e.printStackTrace();
		}
		
	}
	
	public static void main(String[] args) {
		System.out.println("Start:");
		GeneratorInterface ReportGenerator = new GeneratorInterfaceImpl();
		ReportGenerator.startThread();
		System.out.println("End.");
		
	}
}