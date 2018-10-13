package com.staconfree.poi.write.model;

import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月7日
 * @since 1.0
 */
public class FooterModel {
	
	/**
	 * 开始行数
	 */
	public Integer bgnRow;
	
	/**
	 * footer里面的合并区域
	 */
	public List<CellRangeAddress> cellRangeAddresses = new ArrayList<CellRangeAddress>();
	
	public List<TemplateRow> rows = new ArrayList<TemplateRow>();
	
	public List<CellRangeAddress> getCellRangeAddresses() {
		return cellRangeAddresses;
	}
	
	public void setCellRangeAddresses(List<CellRangeAddress> cellRangeAddresses) {
		this.cellRangeAddresses = cellRangeAddresses;
	}
	
	public List<TemplateRow> getRows() {
		return rows;
	}
	
	public void setRows(List<TemplateRow> rows) {
		this.rows = rows;
	}
	
	public Integer getBgnRow() {
		return bgnRow;
	}
	
	public void setBgnRow(Integer bgnRow) {
		this.bgnRow = bgnRow;
	}
	
}
