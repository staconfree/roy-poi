package com.staconfree.poi.write.common;

import com.staconfree.poi.PoiConstants;
import com.staconfree.poi.write.DataLoopModelHelper;
import com.staconfree.poi.write.WriteCellHelper;
import com.staconfree.poi.write.model.DataLoopModel;
import com.staconfree.poi.write.model.TemplateCell;
import com.staconfree.util.JSONUtil;
import net.sf.json.JSONArray;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月4日
 * @since 1.0
 */
public class SimpleDataLoopModelHelper implements DataLoopModelHelper {

	private WriteCellHelper writeCellHelper = new SimpleWriteCellHelper();

	@Override
	public List<DataLoopModel> get(Sheet templatSheet, Workbook srcWorkbook, Workbook desWorkbook) throws Exception {
		List<DataLoopModel> dataLoopModels = new ArrayList<DataLoopModel>();
		int i = 0;
		while (i <= templatSheet.getLastRowNum()) {
			DataLoopModel dataLoopModel = get(templatSheet, i, srcWorkbook, desWorkbook);
			if (dataLoopModel == null) {
				i++;
			} else {
				i = i + dataLoopModel.getRows().size() + 1;
				// 如果前面有多个dataLoopModel，则后面的dataLoopModel的bgnRow需要扣掉前面表达式所占的行数
				for (DataLoopModel model : dataLoopModels) {
					dataLoopModel.setBgnRow(dataLoopModel.getBgnRow() - (2 + model.getRows().size()));
				}
				dataLoopModels.add(dataLoopModel);
			}
		}

		return dataLoopModels;
	}

	/**
	 * 
	 * 所有在bgnInsertRow之后的dataLoopModel的bgnRow都要+insertNum
	 * 
	 * @param dataLoopModels
	 * @param bgnInsertRow
	 *            第几行开始插入
	 * @param insertNum
	 *            插入行数
	 */
	public void updateBgnRow(List<DataLoopModel> dataLoopModels, int bgnInsertRow, int insertNum) throws Exception {
		if (dataLoopModels != null && !dataLoopModels.isEmpty() && bgnInsertRow >= 0 && insertNum > 0) {
			for (DataLoopModel dataLoopModel : dataLoopModels) {
				if (dataLoopModel.getBgnRow() >= bgnInsertRow) {
					dataLoopModel.setBgnRow(dataLoopModel.getBgnRow() + insertNum);
				}
			}
		}
	}

	private DataLoopModel get(Sheet templatSheet, int loopBgnRow, Workbook srcWorkbook, Workbook desWorkbook)
			throws Exception {
		// 配置的格式
		// $$BEGINLOOP{name:"studentList",mergedbaseon:[[0,1],[2,3]],mergedindex:[[0,1],[2]]}
		if (templatSheet != null && srcWorkbook != null && desWorkbook != null) {
			Row row = templatSheet.getRow(loopBgnRow);
			if (row == null) {
				return null;
			}
			String cellValue = writeCellHelper.getCellValue(row.getCell(0));
			if (cellValue.startsWith(PoiConstants.BEGIONLOOP)) {
				DataLoopModel dataLoopModel = new DataLoopModel();
				dataLoopModel.setBgnRow(loopBgnRow);
				dataLoopModel.setTemplateBgnRow(loopBgnRow);
				Map<String, Object> map = JSONUtil
						.conversionToMap(cellValue.substring(PoiConstants.BEGIONLOOP.length()).trim());
				// 数据集名称
				dataLoopModel.setName(map.get(PoiConstants.LOOP_VARNAME).toString());
				// 合并单元格
				// 根据哪些列是否都相同
				JSONArray templateMergedBaseOn = (JSONArray) map.get(PoiConstants.LOOP_MERGEDBASEON);
				// 都相同的时候合并哪一列
				JSONArray templateMergedIndex = (JSONArray) map.get(PoiConstants.LOOP_MERGEDINDEX);
				List<DataLoopModel.Mergedregion> mergedregions = null;
				if (templateMergedBaseOn != null) {
					if (templateMergedIndex == null) {
						templateMergedIndex = templateMergedBaseOn;
					}
					if (templateMergedIndex.size() == templateMergedBaseOn.size()) {
						for (int jsonCur = 0; jsonCur < templateMergedBaseOn.size(); jsonCur++) {
							JSONArray templateMergedBaseOnArr = templateMergedBaseOn.getJSONArray(jsonCur);
							Integer[] mergeBaseOns = (Integer[]) templateMergedBaseOnArr
									.toArray(new Integer[templateMergedBaseOnArr.size()]);
							JSONArray templateMergedIndexArr = templateMergedIndex.getJSONArray(jsonCur);
							Integer[] mergeIndexs = (Integer[]) templateMergedIndexArr
									.toArray(new Integer[templateMergedIndexArr.size()]);
							if (mergedregions == null) {
								mergedregions = new ArrayList<DataLoopModel.Mergedregion>();
							}
							DataLoopModel.Mergedregion mergedregion = new DataLoopModel().new Mergedregion();
							mergedregion.setMergeBaseOn(mergeBaseOns);
							mergedregion.setMergeIndex(mergeIndexs);
							mergedregions.add(mergedregion);
						}
					}
				}
				dataLoopModel.setMergedregions(mergedregions);// celltype是数字的情况
				// 循环体
				int i = loopBgnRow + 1;
				boolean hasEnd = false;
				List<List<TemplateCell>> rowColumns = new ArrayList<List<TemplateCell>>();
				while (i <= templatSheet.getLastRowNum()) {
					row = templatSheet.getRow(i);
					cellValue = writeCellHelper.getCellValue(row.getCell(0));
					if (cellValue.startsWith(PoiConstants.ENDLOOP)) {
						hasEnd = true;
						break;
					}
					List<TemplateCell> columns = new ArrayList<TemplateCell>();
					for (int j = 0; j < row.getLastCellNum(); j++) {
						TemplateCell templateCellStyle = new TemplateCell();
						Cell cell = row.getCell(j);
						cellValue = writeCellHelper.getCellValue(cell);
						if (cellValue == null) {
							cellValue = "";
						}
						templateCellStyle.setCellContent(cellValue);
						templateCellStyle.setCellStyle(writeCellHelper.cloneCellStyle(row.getCell(j).getCellStyle(),
								desWorkbook));
						templateCellStyle.setCellHight(row.getHeight());
						columns.add(templateCellStyle);
					}
					rowColumns.add(columns);
					i++;
				}
				dataLoopModel.setRows(rowColumns);
				if (hasEnd) {
					return dataLoopModel;
				}
			}
		}
		return null;
	}

	public WriteCellHelper getWriteCellHelper() {
		return writeCellHelper;
	}

	public void setWriteCellHelper(WriteCellHelper writeCellHelper) {
		this.writeCellHelper = writeCellHelper;
	}

}
