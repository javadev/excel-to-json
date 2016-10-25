package com.github.javadev.exceltojson.pojo;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;

import org.codehaus.jackson.JsonGenerationException;
import org.codehaus.jackson.map.JsonMappingException;
import org.codehaus.jackson.map.ObjectMapper;

public class ExcelWorkbook {

	private String fileName;
	private Collection<ExcelWorksheet> sheets = new ArrayList<ExcelWorksheet>();
	
	public void addExcelWorksheet(ExcelWorksheet sheet) {
		sheets.add(sheet);
	}
	
	public String toJson(boolean pretty) {
		try {
			if (pretty) {
				return new ObjectMapper().writer().withDefaultPrettyPrinter().writeValueAsString(this);	
			} else {
				return new ObjectMapper().writer().withPrettyPrinter(null).writeValueAsString(this);	
			}
		} catch (JsonGenerationException e) {
			e.printStackTrace();
		} catch (JsonMappingException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	
	public String getFileName() {
		return fileName;
	}

	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	public Collection<ExcelWorksheet> getSheets() {
		return sheets;
	}

	public void setSheets(Collection<ExcelWorksheet> sheets) {
		this.sheets = sheets;
	}

}
