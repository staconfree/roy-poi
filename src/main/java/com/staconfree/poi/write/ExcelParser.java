package com.staconfree.poi.write;

import java.util.Map;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月6日
 * @since 1.0
 */
public interface ExcelParser {
	
	public void setTemplateFilePath(String templateFilePath);
	
	public void setTemplateSheetName(String templateSheetName);
	
	public void setWriteFilePath(String writeFilePath);
	
	public void parse(Map<String, Object> dataStack) throws Exception;
	
	public void endSheet() throws Exception;
	
	public void flush() throws Exception;
}
