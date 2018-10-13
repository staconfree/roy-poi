package com.staconfree.poi.read.xssf;

import com.staconfree.poi.read.ReadCellHelper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月5日
 * @since 1.0
 */
public class XssfReadCellHelper implements ReadCellHelper {
	
	@Override
	public String getCellValue(Cell cell)  throws Exception{
		String value = "";
		if (cell != null) {
			switch (cell.getCellType()) {
				case Cell.CELL_TYPE_FORMULA:
					try {
						value = String.valueOf(cell.getNumericCellValue());
					} catch (IllegalStateException e) {
						value = String.valueOf(cell.getRichStringCellValue());
					}
					break;
				case Cell.CELL_TYPE_NUMERIC:
					// 获取到的单元格类型为数值时做特殊处理，确保表格中的内容和存入的内容一致
					DataFormatter dataFormatter = new DataFormatter();
					value = dataFormatter.formatCellValue(cell);
					// 当数值的长度超过15位时会出现精度问题，取到的值和单元格中的数值不同
					if (!DateUtil.isCellDateFormatted(cell) && value.length() > 15) {
						int leftZeroCount = value.length() - 15;
						String resultString = "";
						// 由于存在精度的问题，需要对第16位数值进行判断，四舍五入
						int sixteenthInt = Integer.parseInt(value.substring(15, 16));
						if (sixteenthInt > 5) {
							int fifteenthInt = Integer.parseInt(value.substring(14, 15)) + 1;
							resultString = value.substring(0, 14) + fifteenthInt;
						} else {
							resultString = value.substring(0, 15);
						}
						
						// 表格中数值的长度大于15位时，大于15位的部分会显示为0
						for (int k = 0; k < leftZeroCount; k++) {
							resultString += "0";
						}
						value = resultString;
					}
					break;
				case Cell.CELL_TYPE_STRING:
					value = String.valueOf(cell.getRichStringCellValue());
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					value = String.valueOf(cell.getBooleanCellValue());
					break;
			}
		}
		return value;
	}
	
}
