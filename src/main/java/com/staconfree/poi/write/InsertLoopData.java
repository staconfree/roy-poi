package com.staconfree.poi.write;

import com.staconfree.poi.write.model.CommonCellStyle;
import com.staconfree.poi.write.model.DataLoopModel;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * 插入数据
 * 
 * @author xiexq
 * @date 2014年3月5日
 * @since 1.0
 */
public interface InsertLoopData {
	/**
	 * 循环插入dataList的数据，并把字段变量转换成string再填入，如果在dataList找不到字段变量，则到dataStack里面找，
	 * 如果都找不到，填空
	 * 
	 * @param sheet
	 * @param dataList
	 * @param dataLoopModel
	 * @param dataStack
	 * @param loopRangeAddresses
	 *            返回插入数据后的合并单元格
	 * @return 返回插入的条数
	 * @throws Exception
	 */
	public int insert(Sheet sheet, Collection<Object> dataList, DataLoopModel dataLoopModel,
                      Map<String, Object> dataStack, List<CellRangeAddress> loopRangeAddresses) throws Exception;

	/**
	 * 只是循环按顺序插入dataList，不涉及字段转换
	 * 
	 * @param sheet
	 * @param dataList
	 * @param dataLoopModel
	 * @param loopRangeAddresses
	 *            返回插入数据后的合并单元格
	 * @return 返回插入的条数
	 * @throws Exception
	 */
	public int insert(Sheet sheet, List<List<String>> dataList, DataLoopModel dataLoopModel,
                      List<CellRangeAddress> loopRangeAddresses) throws Exception;

	/**
	 * 插入一行数据
	 * 
	 * @param sheet
	 * @param dataList
	 * @param dataLoopModel
	 * @return 返回是否成功
	 * @throws Exception
	 */
	public boolean insertOneRow(Sheet sheet, List<String> dataList, DataLoopModel dataLoopModel,
                                List<CellRangeAddress> loopRangeAddresses) throws Exception;

	/**
	 * 
	 * @param workbook
	 * @param sheetName
	 * @param insertBgnRow
	 *            从第几行开始写入
	 * @param rowHight
	 *            列高
	 * @param dataList
	 *            数据集
	 * @param commonCellStyles
	 *            每列的样式
	 * @return
	 * @throws Exception
	 */
	public int simpleInsert(Workbook workbook, String sheetName, int insertBgnRow, short rowHight,
                            List<List<String>> dataList, List<CommonCellStyle> commonCellStyles) throws Exception;

	/**
	 * 规定以什么方式写单元格（拷贝样式，拷贝value，设置样式等）
	 * 
	 * @param writeCellHelper
	 */
	public void setWriteCellHelper(WriteCellHelper writeCellHelper);

}
