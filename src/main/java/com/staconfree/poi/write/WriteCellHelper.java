package com.staconfree.poi.write;

import com.staconfree.poi.write.model.CommonCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月4日
 * @since 1.0
 */
public interface WriteCellHelper {
	/**
	 * 克隆单元格样式
	 * 
	 * @param srcCellStyle
	 * @param desWorkbook
	 * @return
	 * @throws Exception
	 */
	public CellStyle cloneCellStyle(CellStyle srcCellStyle, Workbook desWorkbook) throws Exception;

	/**
	 * 克隆单元格样式
	 * 
	 * @param srcCellStyle
	 * @param desWorkbook
	 * @return
	 * @throws Exception
	 */
	public CellStyle cloneCellStyle(CommonCellStyle srcCellStyle, Workbook desWorkbook) throws Exception;

	/**
	 * 读出，并转为String
	 * 
	 * @param cell
	 * @return
	 */
	public String getCellValue(Cell cell);
	
	/**
	 * 写入，不规定数字格式
	 * 
	 * @param cell
	 * @param value
	 */
	public void setCellValue(Cell cell, String value);
}
