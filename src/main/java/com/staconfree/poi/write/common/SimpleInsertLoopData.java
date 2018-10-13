package com.staconfree.poi.write.common;

import com.staconfree.clazz.reflect.DataReflect;
import com.staconfree.exception.ServiceLevelException;
import com.staconfree.poi.write.DataLoopMergedRegion;
import com.staconfree.poi.write.InsertLoopData;
import com.staconfree.poi.write.WriteCallBack;
import com.staconfree.poi.write.WriteCellHelper;
import com.staconfree.poi.write.model.CommonCellStyle;
import com.staconfree.poi.write.model.DataLoopModel;
import com.staconfree.poi.write.model.TemplateCell;
import com.staconfree.var.VarConverter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.*;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月5日
 * @since 1.0
 */
public class SimpleInsertLoopData implements InsertLoopData {

	private WriteCellHelper writeCellHelper = new SimpleWriteCellHelper();

	private WriteCallBack writeCallBack;

	public int insert(Sheet sheet, Collection<Object> dataList, DataLoopModel dataLoopModel,
			Map<String, Object> dataStack, List<CellRangeAddress> loopRangeAddresses) throws Exception {
		if (sheet != null && dataList != null && !dataList.isEmpty() && dataLoopModel != null) {
			int templateRowNum = dataLoopModel.getRows().size();
			Class clazz = dataList.iterator().next().getClass();
			DataReflect dataReflect = new DataReflect(clazz);
			List<List<TemplateCell>> templateLoopRowColums = dataLoopModel.getRows();
			int writeRowIndex = dataLoopModel.getBgnRow();
			Iterator<Object> it = dataList.iterator();
			int dataIndex = 0;
			while (it.hasNext()) {
				Object data = it.next();
				DataLoopMergedRegion dataLoopMergedRegion = new DataLoopMergedRegion();
				for (int datatemplateIndex = 0; datatemplateIndex < templateRowNum; datatemplateIndex++) {// 行写入
					Row newrow = (Row) sheet.createRow(writeRowIndex);
					// 设置高度
					newrow.setHeight(templateLoopRowColums.get(0).get(0).getCellHight());
					List<TemplateCell> columns = templateLoopRowColums.get(datatemplateIndex);
					int columnSize = columns.size();
					for (int columnIndex = 0; columnIndex < columnSize; columnIndex++) {// 列写入
						TemplateCell templateCellStyle = columns.get(columnIndex);
						Cell newcell = newrow.createCell(columnIndex);
						// 数据转换
						String formula = templateCellStyle.getCellContent();
						String changeValue = VarConverter.convertLoopData(formula, data, dataReflect, dataStack,
                                dataIndex, dataLoopModel.getLoopedNum());
						// 设值
						if (changeValue != null) {
							changeValue = VarConverter.convertSystemData(changeValue);
							writeCellHelper.setCellValue(newcell, changeValue);
						}
						// 设样式
//						CellStyle newCellStyle = writeCellHelper.cloneCellStyle(templateCellStyle.getCellStyle(), sheet.getWorkbook());
						CellStyle newCellStyle = templateCellStyle.getCellStyle();
						newcell.setCellStyle(newCellStyle);
//						System.out.println("row="+newrow.getRowNum() +"," +"colum="+ newcell.getColumnIndex() +","+templateCellStyle.getIndex()+","+newCellStyle.getBorderRight());
//						if (newrow.getRowNum() == 3640) {
//							System.out.println("格式要丢了");
//						}
						// 合并单元格
						if (templateRowNum == 1) {// 只有循环体为一行的时候才会进行合并判断
							dataLoopMergedRegion.reMerged(loopRangeAddresses, writeRowIndex, columnIndex, columnSize,
									changeValue, dataLoopModel);
						}
						if (writeCallBack != null) {
							writeCallBack.changeCell(newcell, data, sheet);
						}
					}
					if (writeCallBack != null) {
						writeCallBack.changeRow(newrow, data, sheet);
					}
					// 写入合并单元格
					dataLoopMergedRegion.writeMerged(sheet, writeRowIndex, loopRangeAddresses);
					writeRowIndex++; // for循环里面才加
				}

				dataIndex++;
			}
			int insertNum = writeRowIndex - dataLoopModel.getBgnRow();
			dataLoopModel.setLoopedNum(dataLoopModel.getLoopedNum() + insertNum);
			return insertNum;
		}
		return 0;
	}

	public int insert(Sheet sheet, List<List<String>> dataList, DataLoopModel dataLoopModel,
			List<CellRangeAddress> loopRangeAddresses) throws Exception {
		if (sheet != null && dataList != null && !dataList.isEmpty() && dataLoopModel != null) {
			int dataNum = dataList.size();
			for (int dataIndex = 0; dataIndex < dataNum; dataIndex++) {
				List<String> data = dataList.get(dataIndex);
				if (data.size() != dataLoopModel.getRows().get(0).size()) {
					throw new ServiceLevelException(0, "数据与样式的列数不等！");
				}
				// 一行行插入
				insertOneRow(sheet, data, dataLoopModel, loopRangeAddresses);
			}
			return dataNum;
		}
		return 0;
	}

	public boolean insertOneRow(Sheet sheet, List<String> dataList, DataLoopModel dataLoopModel,
			List<CellRangeAddress> loopRangeAddresses) throws Exception {
		List<TemplateCell> templateCells = dataLoopModel.getRows().get(0);
		int writeRowIndex = dataLoopModel.getBgnRow();
		Row newrow = sheet.createRow(dataLoopModel.getBgnRow());
		// 设置高度
		newrow.setHeight(dataLoopModel.getRowHight());
		int columnSize = dataList.size();
		// 合并单元格
		DataLoopMergedRegion dataLoopMergedRegion = new DataLoopMergedRegion();
		for (int columnIndex = 0; columnIndex < columnSize; columnIndex++) {
			// 列写入
			Cell newcell = newrow.createCell(columnIndex);
			// 数据转换
			String cellValue = dataList.get(columnIndex);
			// 设值
			if (cellValue != null) {
				writeCellHelper.setCellValue(newcell, cellValue);
			}
			// 设样式
			if (templateCells != null && templateCells.get(columnIndex) != null) {
				newcell.setCellStyle(templateCells.get(columnIndex).getCellStyle());
			}
			if (writeCallBack != null) {
				writeCallBack.changeCell(newcell,dataList,sheet);
			}
			dataLoopMergedRegion.reMerged(loopRangeAddresses, writeRowIndex, columnIndex, columnSize, cellValue,
					dataLoopModel);
		}
		if (writeCallBack != null) {
			writeCallBack.changeRow(newrow,dataList,sheet);
		}
		// 写入合并单元格
		dataLoopMergedRegion.writeMerged(sheet, writeRowIndex, loopRangeAddresses);
		dataLoopModel.setBgnRow(dataLoopModel.getBgnRow() + 1);
		return false;
	}

	public int simpleInsert(Workbook workbook, String sheetName, int insertBgnRow, short rowHight,
			List<List<String>> dataList, List<CommonCellStyle> commonCellStyles) throws Exception {
		if (sheetName != null && dataList != null && !dataList.isEmpty()) {
			Sheet sheet = workbook.getSheet(sheetName);
			DataLoopModel dataLoopModel = new DataLoopModel();
			dataLoopModel.setBgnRow(insertBgnRow);
			dataLoopModel.setRowHight(rowHight);
			List<TemplateCell> row = new ArrayList<TemplateCell>();
			for (int i = 0; i < commonCellStyles.size(); i++) {
				CellStyle cellStyle = writeCellHelper.cloneCellStyle(commonCellStyles.get(i), workbook);
				TemplateCell cell = new TemplateCell();
				cell.setCellStyle(cellStyle);
				row.add(cell);
			}
			List<List<TemplateCell>> rows = new ArrayList<List<TemplateCell>>();
			rows.add(row);
			dataLoopModel.setRows(rows);
			this.insert(sheet, dataList, dataLoopModel, null);
		}
		return 0;
	}

	public WriteCellHelper getWriteCellHelper() {
		return writeCellHelper;
	}

	public void setWriteCellHelper(WriteCellHelper writeCellHelper) {
		this.writeCellHelper = writeCellHelper;
	}

	public WriteCallBack getWriteCallBack() {
		return writeCallBack;
	}

	public void setWriteCallBack(WriteCallBack writeCallBack) {
		this.writeCallBack = writeCallBack;
	}

}
