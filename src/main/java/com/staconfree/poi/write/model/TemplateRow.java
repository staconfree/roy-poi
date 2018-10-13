package com.staconfree.poi.write.model;

import java.util.ArrayList;
import java.util.List;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月10日
 * @since 1.0
 */
public class TemplateRow {
	/**
	 * 行高
	 */
	public short rowHight;
	/**
	 * 列
	 */
	public List<TemplateCell> cells = new ArrayList<TemplateCell>();
	
	public short getRowHight() {
		return rowHight;
	}
	
	public void setRowHight(short rowHight) {
		this.rowHight = rowHight;
	}
	
	public List<TemplateCell> getCells() {
		return cells;
	}
	
	public void setCells(List<TemplateCell> cells) {
		this.cells = cells;
	}
	
}
