package services;

import java.io.IOException;


public interface GeneratorInterface {
	public abstract void startThread();
	public String processExportExcel(String sourceDirectory) throws IOException;
}