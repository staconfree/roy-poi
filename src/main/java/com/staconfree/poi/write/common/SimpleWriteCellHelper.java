package com.staconfree.poi.write.common;

import com.staconfree.poi.write.WriteCellHelper;
import com.staconfree.poi.write.model.CommonCellStyle;
import org.apache.poi.ss.usermodel.*;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 2007版excel的写单元格类
 * 
 * @author xiexq
 * @date 2014年3月4日
 * @since 1.0
 */
public class SimpleWriteCellHelper implements WriteCellHelper {

	private static Pattern pattern = Pattern.compile("0|-?0.[0-9]+|-?[1-9][0-9]*.?[0-9]*");
	
	@Override
	public CellStyle cloneCellStyle(CellStyle srcCellStyle, Workbook desWorkbook) throws Exception {
		if (srcCellStyle != null && desWorkbook != null) {
			CellStyle desCellStyle = desWorkbook.createCellStyle();
			// 拷贝其他通用属性
			cloneCellStyle(srcCellStyle, desCellStyle, desWorkbook);
			return desCellStyle;
		}
		return null;
	}

	public CellStyle cloneCellStyle(CommonCellStyle srcCellStyle, Workbook desWorkbook) throws Exception {
		if (desWorkbook != null && srcCellStyle != null) {
			CellStyle desCellStyle = desWorkbook.createCellStyle();
			// 拷贝其他通用属性
			cloneCellStyle(srcCellStyle, desCellStyle, desWorkbook);
			return desCellStyle;
		}
		return null;
	}

	protected void cloneCellStyle(CommonCellStyle srcCellStyle, CellStyle desCellStyle, Workbook desWorkbook) {
		if (srcCellStyle != null && desCellStyle != null && desWorkbook != null) {
			// 设置字体
			Font newFont = desWorkbook.createFont();
			newFont.setFontHeightInPoints(srcCellStyle.getFontHeightInPoints());
			newFont.setFontName(srcCellStyle.getFontName());
			newFont.setBoldweight(srcCellStyle.getFontBoldweight());
			desCellStyle.setFont(newFont);

			// 设置边框线型
			// 注：这里不能设置边框颜色，否则会导致导出的单元格不能设置单元格格式
			desCellStyle.setBorderTop(srcCellStyle.getBorderTop());
			desCellStyle.setBorderBottom(srcCellStyle.getBorderBottom());
			desCellStyle.setBorderLeft(srcCellStyle.getBorderLeft());
			desCellStyle.setBorderRight(srcCellStyle.getBorderRight());
			// 设置内容位置:例水平居中,居右，居工
			desCellStyle.setAlignment(srcCellStyle.getAlignment());
			// 设置内容位置:例垂直居中,居上，居下
			desCellStyle.setVerticalAlignment(srcCellStyle.getVerticalAlignment());
			// 自动换行
			desCellStyle.setWrapText(srcCellStyle.isWrapText());
			// 设置背景颜色
			// 其他属性
			desCellStyle.setDataFormat(srcCellStyle.getDataFormat());
			desCellStyle.setHidden(srcCellStyle.isHidden());
			desCellStyle.setIndention(srcCellStyle.getIndention());// 首行缩进
			desCellStyle.setLocked(srcCellStyle.isLocked());
			desCellStyle.setRotation(srcCellStyle.getRotation());// 旋转
		}
	}

	protected void cloneCellStyle(CellStyle srcCellStyle, CellStyle desCellStyle, Workbook desWorkbook)
			throws Exception {
		if (srcCellStyle != null && desCellStyle != null && desWorkbook != null) {
			// 设置字体
			Font newFont = desWorkbook.createFont();
			newFont.setFontHeightInPoints(CommonCellStyle.FONTHEIGHTINPOINTS);
			newFont.setFontName(CommonCellStyle.FONTNAME);
			newFont.setBoldweight(CommonCellStyle.FONTBOLDWEIGHT);
			desCellStyle.setFont(newFont);
			// 设置边框线型
			desCellStyle.setBorderTop(srcCellStyle.getBorderTop());
			desCellStyle.setBorderBottom(srcCellStyle.getBorderBottom());
			desCellStyle.setBorderLeft(srcCellStyle.getBorderLeft());
			desCellStyle.setBorderRight(srcCellStyle.getBorderRight());
			// 设置内容位置:例水平居中,居右，居工
			desCellStyle.setAlignment(srcCellStyle.getAlignment());
			// 设置内容位置:例垂直居中,居上，居下
			desCellStyle.setVerticalAlignment(srcCellStyle.getVerticalAlignment());
			// 自动换行
			desCellStyle.setWrapText(srcCellStyle.getWrapText());

			// 其他属性
			desCellStyle.setDataFormat(srcCellStyle.getDataFormat());
			desCellStyle.setHidden(srcCellStyle.getHidden());
			desCellStyle.setIndention(srcCellStyle.getIndention());// 首行缩进
			desCellStyle.setLocked(srcCellStyle.getLocked());
			desCellStyle.setRotation(srcCellStyle.getRotation());// 旋转
			// 设置背景颜色
			desCellStyle.setFillPattern(CellStyle.NO_FILL);
		}
	}

	@Override
	public String getCellValue(Cell cell) {
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
	
	public void setCellValue(Cell cell, String value) {
		if (cell != null && value != null) {
			if (isDigit(value)) {
				try {
					Double d = Double.valueOf(value);
					if (d > 100000000000.0) {// 如果数值太大，还是使用String格式，否则显示出来会变成科学计数法
						cell.setCellType(Cell.CELL_TYPE_STRING);
						cell.setCellValue(value);
					} else {
						cell.setCellType(Cell.CELL_TYPE_NUMERIC);
						cell.setCellValue(d);
					}
				} catch (Exception e) {
					cell.setCellType(Cell.CELL_TYPE_STRING);
					cell.setCellValue(value);
				}
			} else {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(value);
			}
		}
	}
	
	private boolean isDigit(String value) {
		Matcher matcher = pattern.matcher(value);
		return matcher.matches();
	}
	
	public static void main(String[] args) {
		digit("-404");
		digit("404");
		digit("0");
		digit("0.11");
		digit("1.11");
		digit("2280.11");
		digit("2280.0000");
		digit("-65.6667");
	}
	
	public static void digit(String value) {
		Matcher matcher = pattern.matcher(value);
		if (matcher.matches()) {
			System.out.println(value);
		} else {
			System.out.println("not found");
		}
	}
	
}
