package com.staconfree.poi.write;

import com.staconfree.poi.MergedRegionUtil;
import com.staconfree.poi.write.model.DataLoopModel;
import com.staconfree.util.JSONUtil;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 处理数据循环区域的合并单元格
 * 
 * @author xiexq
 * @date 2014年3月5日
 * @since 1.0
 */
public class DataLoopMergedRegion {

	/**
	 * 写入[writeRowIndex,writeColumnIndex]数据的时候，通过dataLoopModel里面记录的值,
	 * 判断是否需要延长合并单元格
	 * 
	 * @param loopRangeAddresses
	 * @param writeRowIndex
	 * @param writeColumnIndex
	 * @param columnSize
	 * @param cellValue
	 * @param dataLoopModel
	 */
	public void reMerged(List<CellRangeAddress> loopRangeAddresses, int writeRowIndex, int writeColumnIndex,
			int columnSize, String cellValue, DataLoopModel dataLoopModel) throws Exception {
		List<DataLoopModel.Mergedregion> mergedregions = dataLoopModel.getMergedregions();
		if (mergedregions == null || mergedregions.isEmpty()) {
			return;
		}
		for (DataLoopModel.Mergedregion mergedregion : mergedregions) {
			if (contains(mergedregion.getMergeBaseOn(), writeColumnIndex)) {
				Map<String, String> thisValueMap = new HashMap<String, String>();
				String thisRowValue = mergedregion.getThisRowValue();
				if (thisRowValue != null && !"".equals(thisRowValue)) {
					thisValueMap = JSONUtil.conversionToMap(mergedregion.getThisRowValue());
				}
				thisValueMap.put(String.valueOf(writeColumnIndex), cellValue);
				mergedregion.setThisRowValue(JSONUtil.conversionToJSON(thisValueMap));
			}
		}
		if (writeColumnIndex == columnSize - 1) {// 到了最后一列，进行合并单元格
			for (DataLoopModel.Mergedregion mergedregion : mergedregions) {
				String thisRowValue = mergedregion.getThisRowValue();
				String lastRowValue = mergedregion.getLastRowValue();
				if (!"".equals(thisRowValue)&&thisRowValue.equals(lastRowValue)) {// 如果前后两次都相同，则合并mergeIndex
					Integer[] mergeIndexs = mergedregion.getMergeIndex();
					for (Integer mergedColIndex : mergeIndexs) {
						Map<String, String> thisValueMap = JSONUtil.conversionToMap(mergedregion.getThisRowValue());
						String thisCellValue = thisValueMap.get(mergedColIndex+"");
						if (!"".equals(thisCellValue)){
							CellRangeAddress lastCellRangeAddress = MergedRegionUtil.getCellRangeAddress(
									loopRangeAddresses, writeRowIndex - 1, mergedColIndex);
							// 先删除上次的合并区域
							MergedRegionUtil.removeCellRangeAddress(loopRangeAddresses, lastCellRangeAddress);
							lastCellRangeAddress.setLastRow(writeRowIndex);
							// 再加入最新的合并区域
							loopRangeAddresses.add(lastCellRangeAddress);
						}
					}
				}
				mergedregion.setLastRowValue(thisRowValue);
			}
		}
	}

	/**
	 * 把loopRangeAddresses里面的合并单元格在writeRowIndex之前结束的，写入sheet里面，
	 * 同时删掉loopRangeAddresses里面已写入的,防止loopRangeAddresses一直膨胀下去
	 * 
	 * @param sheet
	 * @param writeRowIndex
	 * @param loopRangeAddresses
	 * @throws Exception
	 */
	public void writeMerged(Sheet sheet, int writeRowIndex, List<CellRangeAddress> loopRangeAddresses) throws Exception {
		if (loopRangeAddresses != null && !loopRangeAddresses.isEmpty() && sheet != null) {
			int size = loopRangeAddresses.size();
			for (int i = size - 1; i >= 0; i--) {
				CellRangeAddress cellRangeAddress = loopRangeAddresses.get(i);
				if (cellRangeAddress.getLastRow() < writeRowIndex) {
					sheet.addMergedRegion(cellRangeAddress);
					loopRangeAddresses.remove(cellRangeAddress);
				}
			}
		}
	}

	private boolean contains(Integer[] intArr, int param) {
		if (intArr != null) {
			for (Integer integer : intArr) {
				if (integer.intValue() == param) {
					return true;
				}
			}
		}
		return false;
	}

}
