package com.staconfree.poi.write;

import com.staconfree.poi.write.model.FooterModel;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 插入文件尾
 * 
 * @author xiexq
 * @date 2014年3月10日
 * @since 1.0
 */
public interface InsertFooter {
	/**
	 * 
	 * @param sheet
	 * @param footerModel
	 * @param insertLoopNum
	 *            插入的循环数据的行数
	 * @throws Exception 
	 */
	public void insert(Workbook workbook, Sheet sheet, FooterModel footerModel, int insertLoopNum) throws Exception;
	
	/**
	 * 规定以什么方式写单元格（拷贝样式，拷贝value，设置样式等）
	 * 
	 * @param writeCellHelper
	 */
	public void setWriteCellHelper(WriteCellHelper writeCellHelper);
	
}
