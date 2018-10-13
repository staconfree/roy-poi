package com.staconfree.poi.write;

import com.staconfree.poi.write.model.DataLoopModel;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月4日
 * @since 1.0
 */
public interface DataLoopModelHelper {
	/**
	 * 
	 * @param templatSheet
	 * @return
	 * @throws Exception 
	 */
	public List<DataLoopModel> get(Sheet templatSheet, Workbook desWorkbook, Workbook desWorkBook) throws Exception;
	
	/**
	 * 
	 * 所有在bgnInsertRow之后的dataLoopModel的bgnRow都要+insertNum
	 * 
	 * @param dataLoopModels
	 * @param bgnInsertRow
	 *            第几行开始插入
	 * @param insertNum
	 *            插入行数
	 * @throws Exception 
	 */
	public void updateBgnRow(List<DataLoopModel> dataLoopModels, int bgnInsertRow, int insertNum) throws Exception;
	
	/**
	 * 规定以什么方式写单元格（拷贝样式，拷贝value，设置样式等）
	 * 
	 * @param writeCellHelper
	 */
	public void setWriteCellHelper(WriteCellHelper writeCellHelper);	
}
