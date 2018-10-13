package com.staconfree.poi.write.common;

import com.staconfree.poi.MergedRegionUtil;
import com.staconfree.poi.write.CopyWorkBook;
import com.staconfree.poi.write.WriteCallBack;
import com.staconfree.poi.write.WriteCellHelper;
import com.staconfree.poi.write.model.DataLoopModel;
import com.staconfree.poi.write.model.FooterModel;
import com.staconfree.poi.write.model.TemplateCell;
import com.staconfree.poi.write.model.TemplateRow;
import com.staconfree.var.VarConverter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月5日
 * @since 1.0
 */
public class SimpleCopyWorkBook implements CopyWorkBook {

	/**
	 * 控制50列的列宽跟模板一致，可自己设置
	 */
	private int maxColumns = 50;

	private WriteCellHelper writeCellHelper = new SimpleWriteCellHelper();
	
	private WriteCallBack writeCallBack;

	public void copy(Workbook srcWorkbook, String srcSheetName, Workbook desWorkbook, String desSheetName,
			List<DataLoopModel> dataLoopModels, FooterModel footerModel, Map<String, Object> dataStack)
			throws Exception {
		if (srcWorkbook != null && desWorkbook != null && dataLoopModels != null && footerModel != null) {
			Sheet srcSheet = srcWorkbook.getSheet(srcSheetName);
			Sheet desSheet = desWorkbook.createSheet(desSheetName);
			// 设置列宽
			for (int i = 0; i < maxColumns; i++) {
				int columnwidth = srcSheet.getColumnWidth(i);
				desSheet.setColumnWidth(i, columnwidth);
			}
			// 获取数据循环配置
			SimpleDataLoopModelHelper dataLoopModelHelper = new SimpleDataLoopModelHelper();
			dataLoopModelHelper.setWriteCellHelper(writeCellHelper);
			List<DataLoopModel> dataLoopModels_tmp = dataLoopModelHelper.get(srcSheet, srcWorkbook, desWorkbook);
			dataLoopModels.addAll(dataLoopModels_tmp);
			// 获取合并区域
			List<CellRangeAddress> regions = new ArrayList<CellRangeAddress>();
			for (int i = 0; i < srcSheet.getNumMergedRegions(); i++) {
				CellRangeAddress region = srcSheet.getMergedRegion(i);
				regions.add(region);
			}
			// 复制数据
			int totalRowCount = srcSheet.getLastRowNum();
			int templateRowIndex = 0;
			int writeRowIndex = 0;
			int removeRowNum = 0;
			boolean isFooter = true;
			while (templateRowIndex <= totalRowCount) {
				isFooter = true;
				Row row = srcSheet.getRow(templateRowIndex);
				Row newrow = null;
				DataLoopModel dataLoopMode = null;// 是否开始循环，非空代表开始循环
				for (DataLoopModel model : dataLoopModels) {
					if (model.getTemplateBgnRow() == templateRowIndex) {
						dataLoopMode = model;
					}
					if (model.getTemplateBgnRow() > templateRowIndex) {
						isFooter = false;
					}
				}
				if (dataLoopMode == null) {
					if (!isFooter) {// 写入循环体之外的东西
						newrow = desSheet.createRow(writeRowIndex);
						if (row == null) {// 如果为空，跳过该行
							templateRowIndex++;
							writeRowIndex++;
							continue;
						}
						// 设置高度
						newrow.setHeight(row.getHeight());
						for (int j = 0; j < row.getLastCellNum(); j++) {
							Cell cell = row.getCell(j);
							if (cell == null) {
								continue;
							}
							Cell newcell = newrow.createCell(j);
							CellStyle cellStyle = cell.getCellStyle();
							// 设置单元格样式
							CellStyle newCellStyle = writeCellHelper.cloneCellStyle(cellStyle, desWorkbook);
							newcell.setCellStyle(newCellStyle);
							// 设置单元格的值
							String cellValue = writeCellHelper.getCellValue(cell);
							cellValue = VarConverter.convertCommonData(cellValue, dataStack);
							if (cellValue != null && !"".equals(cellValue)) {
								cellValue = VarConverter.convertSystemData(cellValue);
							}
							writeCellHelper.setCellValue(newcell, cellValue);
							if (writeCallBack != null) {
								writeCallBack.changeCell(newcell, null, null);
							}
						}
						if (writeCallBack != null) {
							writeCallBack.changeRow(newrow, null, null);
						}
						writeRowIndex++;
					} else {// 保存footer
						if (footerModel.getBgnRow() == null) {
							footerModel.setBgnRow(templateRowIndex - removeRowNum);
						}
						TemplateRow footRow = new TemplateRow();
						if (row == null) {// 如果为空，跳过该行
							templateRowIndex++;
							footRow.setRowHight((short) 270);
							footerModel.getRows().add(footRow);
							continue;
						}
						footRow.setRowHight(row.getHeight());
						for (int j = 0; j < row.getLastCellNum(); j++) {
							Cell cell = row.getCell(j);
							if (cell == null) {
								continue;
							}
							CellStyle cellStyle = cell.getCellStyle();
							// 设置单元格样式
							CellStyle newCellStyle = writeCellHelper.cloneCellStyle(cellStyle, srcWorkbook);
							TemplateCell templateCell = new TemplateCell();
							templateCell.setIndex(j);
							templateCell.setCellStyle(newCellStyle);
							// 设置单元格的值
							String cellValue = writeCellHelper.getCellValue(cell);
							cellValue = VarConverter.convertCommonData(cellValue, dataStack);
							if (cellValue != null && !"".equals(cellValue)) {
								cellValue = VarConverter.convertSystemData(cellValue);
							}
							templateCell.setCellContent(cellValue);
							footRow.getCells().add(templateCell);
						}
						footerModel.getRows().add(footRow);
					}
					templateRowIndex++;
				} else {// 跳过循环数据区域
					// 重新调整合并区
					int bgnRemoveRow = templateRowIndex - removeRowNum;
					int endRemoveRow = bgnRemoveRow + dataLoopMode.getRows().size() + 2 - 1;
					MergedRegionUtil.resetByRemoveRows(regions, bgnRemoveRow, endRemoveRow);
					removeRowNum += dataLoopMode.getRows().size() + 2;
					// 跳过循环数据区域
					templateRowIndex = templateRowIndex + dataLoopMode.getRows().size() + 2;
				}
			}
			// 设置合并区域
			for (CellRangeAddress region : regions) {
				if (footerModel.getBgnRow() == null || footerModel.getBgnRow() != null
						&& region.getLastRow() < footerModel.getBgnRow()) {
					desSheet.addMergedRegion(region);
				} else {
					footerModel.getCellRangeAddresses().add(region);
				}
			}
		}
	}

	public WriteCellHelper getWriteCellHelper() {
		return writeCellHelper;
	}

	public void setWriteCellHelper(WriteCellHelper writeCellHelper) {
		this.writeCellHelper = writeCellHelper;
	}

	public int getMaxColumns() {
		return maxColumns;
	}

	public void setMaxColumns(int maxColumns) {
		this.maxColumns = maxColumns;
	}

	public WriteCallBack getWriteCallBack() {
		return writeCallBack;
	}

	public void setWriteCallBack(WriteCallBack writeCallBack) {
		this.writeCallBack = writeCallBack;
	}

}
