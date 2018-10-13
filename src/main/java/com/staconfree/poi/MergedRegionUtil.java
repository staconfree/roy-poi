package com.staconfree.poi;

import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月5日
 * @since 1.0
 */
public class MergedRegionUtil {
	
	/**
	 * 插入几行之后，需要重新设置被影响的合并区域
	 * 
	 * @param cellRangeAddresses
	 *            原先的合并情况
	 * @param bgnRemoveRow
	 *            起始删除的行号
	 * @param endRemoveRow
	 *            截止删除的行号
	 * @return
	 */
	public static void resetByRemoveRows(List<CellRangeAddress> cellRangeAddresses, int bgnRemoveRow, int endRemoveRow)  throws Exception{
		if (cellRangeAddresses != null && !cellRangeAddresses.isEmpty() && bgnRemoveRow >= 0 && endRemoveRow >= bgnRemoveRow) {
			for (CellRangeAddress cellRangeAddress : cellRangeAddresses) {
				int firstRow = cellRangeAddress.getFirstRow();
				int lastRow = cellRangeAddress.getLastRow();
				if (firstRow > endRemoveRow) {
					cellRangeAddress.setFirstRow(firstRow - (endRemoveRow - bgnRemoveRow + 1));
					cellRangeAddress.setLastRow(lastRow - (endRemoveRow - bgnRemoveRow + 1));
				}
			}
		}
	}
	
	/**
	 * 得到某个单元格所在的合并区域，如果没有，就默认单元格本身就是一个合并区域
	 * 
	 * @param sheet
	 * @param cellRow
	 * @param cellCol
	 * @return
	 */
	public static CellRangeAddress getCellRangeAddress(List<CellRangeAddress> loopRangeAddresses, int cellRow, int cellCol) {
		for (CellRangeAddress region : loopRangeAddresses) {
			if (region.getFirstRow() <= cellRow && region.getLastRow() >= cellRow && region.getFirstColumn() <= cellCol
					&& region.getLastColumn() >= cellCol) {
				return region;
			}
		}
		return new CellRangeAddress(cellRow, cellRow, cellCol, cellCol);
	}
	
	/**
	 * 删除指定合并区域
	 * 
	 * @param sheet
	 * @param cellRangeAddress
	 */
	public static void removeCellRangeAddress(List<CellRangeAddress> loopRangeAddresses, CellRangeAddress cellRangeAddress) {
		for (CellRangeAddress region : loopRangeAddresses) {
			if (region.getFirstColumn() == cellRangeAddress.getFirstColumn()
					&& region.getLastColumn() == cellRangeAddress.getLastColumn()
					&& region.getFirstRow() == cellRangeAddress.getFirstRow()
					&& region.getLastRow() == cellRangeAddress.getLastRow()) {
				loopRangeAddresses.remove(region);
				break;
			}
		}
	}
	
}
