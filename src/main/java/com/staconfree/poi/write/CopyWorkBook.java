package com.staconfree.poi.write;

import com.staconfree.poi.write.model.DataLoopModel;
import com.staconfree.poi.write.model.FooterModel;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;
import java.util.Map;

/**
 * 拷贝模板
 * 
 * @author xiexq
 * @date 2014年3月5日
 * @since 1.0
 */
public interface CopyWorkBook {
	/**
	 * 
	 * @param srcWorkbook
	 *            非空
	 * @param srcSheetName
	 * @param desWorkbook
	 *            非空
	 * @param desSheetName
	 * @param dataLoopModels
	 *            非空
	 * @param footerModel
	 *            非空
	 * @param dataStack
	 *            数据堆栈，用于拷贝的时候同时修改一些变量值
	 * @throws Exception 
	 */
	public void copy(Workbook srcWorkbook, String srcSheetName, Workbook desWorkbook, String desSheetName,
                     List<DataLoopModel> dataLoopModels, FooterModel footerModel, Map<String, Object> dataStack) throws Exception;
	
	/**
	 * 规定以什么方式写单元格（拷贝样式，拷贝value，设置样式等）
	 * 
	 * @param writeCellHelper
	 */
	public void setWriteCellHelper(WriteCellHelper writeCellHelper);
}
