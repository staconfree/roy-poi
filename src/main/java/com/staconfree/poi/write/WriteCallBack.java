package com.staconfree.poi.write;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * 
 * @author xiexq
 * @date 2014年4月25日
 * @since 1.0
 */
public interface WriteCallBack {
	/**
	 * 每次写完一行的时候，可以自己再个性化行的内容、样式等
	 * 
	 * @param row
	 * @param rowData 行数据
	 */
	void changeRow(Row row, Object rowData, Sheet sheet);

	/**
	 * 每次写完一个单元格的时候，可以自己再单元格行的内容、样式等
	 *
	 * @param cell
	 * @param rowData 行数据
	 */
	void changeCell(Cell cell, Object rowData, Sheet sheet);

}
