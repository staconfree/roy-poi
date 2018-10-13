package com.staconfree.poi.write.model;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月10日
 * @since 1.0
 */
public class TemplateCell {
	/**
	 * 列的位置
	 */
	private int index;
	/**
	 * 模板列的内容
	 */
	private String cellContent;
	/**
	 * destWorkBook列的样式(用于insert数据的时候使用)
	 */
	private CellStyle cellStyle;
	/**
	 * 列高度
	 */
	private short cellHight;

	public int getIndex() {
		return index;
	}

	public void setIndex(int index) {
		this.index = index;
	}

	public String getCellContent() {
		return cellContent;
	}

	public void setCellContent(String cellContent) {
		this.cellContent = cellContent;
	}

	public CellStyle getCellStyle() {
		return cellStyle;
	}

	public void setCellStyle(CellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}

	public short getCellHight() {
		return cellHight;
	}

	public void setCellHight(short cellHight) {
		this.cellHight = cellHight;
	}
	
}
