package com.staconfree.poi.write.common;

import com.staconfree.poi.write.InsertFooter;
import com.staconfree.poi.write.WriteCallBack;
import com.staconfree.poi.write.WriteCellHelper;
import com.staconfree.poi.write.model.FooterModel;
import com.staconfree.poi.write.model.TemplateCell;
import com.staconfree.poi.write.model.TemplateRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月10日
 * @since 1.0
 */
public class SimpleInsertFooter implements InsertFooter {

	private WriteCellHelper writeCellHelper = new SimpleWriteCellHelper();

	private WriteCallBack writeCallBack;
	
	@Override
	public void insert(Workbook workbook, Sheet sheet, FooterModel footerModel, int insertLoopNum) throws Exception {
		if (workbook == null || sheet == null || footerModel == null || footerModel.getBgnRow() == null) {
			return;
		}
		List<TemplateRow> footRows = footerModel.getRows();
		int bgnRow = footerModel.getBgnRow() + insertLoopNum;
		for (TemplateRow templateRow : footRows) {
			Row newrow = sheet.createRow(bgnRow);
			// 设置高度
			newrow.setHeight(templateRow.getRowHight());
			List<TemplateCell> footCells = templateRow.getCells();
			for (TemplateCell templateCell : footCells) {
				Cell newCell = newrow.createCell(templateCell.getIndex());
				// 设值
				writeCellHelper.setCellValue(newCell, templateCell.getCellContent());
				// 设样式
				newCell.setCellStyle(writeCellHelper.cloneCellStyle(templateCell.getCellStyle(), workbook));
				if (writeCallBack != null) {
					writeCallBack.changeCell(newCell,footCells,sheet);
				}
			}
			if (writeCallBack != null) {
				writeCallBack.changeRow(newrow,footCells,sheet);
			}
			bgnRow++;
		}
		// 重新调整footer的合并区域
		List<CellRangeAddress> footerRangeAddresses = footerModel.getCellRangeAddresses();
		for (CellRangeAddress cellRangeAddress : footerRangeAddresses) {
			cellRangeAddress.setFirstRow(cellRangeAddress.getFirstRow() + insertLoopNum);
			cellRangeAddress.setLastRow(cellRangeAddress.getLastRow() + insertLoopNum);
		}
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
