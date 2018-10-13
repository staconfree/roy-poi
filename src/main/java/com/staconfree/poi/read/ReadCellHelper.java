package com.staconfree.poi.read;

import org.apache.poi.ss.usermodel.Cell;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月4日
 * @since 1.0
 */
public interface ReadCellHelper {
	
	/**
	 * 得到单元格的东西，并转为String
	 * 
	 * @param cell
	 * @return
	 */
	public String getCellValue(Cell cell) throws Exception;
	
}
