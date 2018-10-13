package com.staconfree.poi.write.sxssf;

import com.staconfree.poi.write.common.SimpleWriteCellHelper;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.Map;

/**
 * 2007版excel的写单元格类
 * 
 * @author xiexq
 * @date 2014年3月4日
 * @since 1.0
 */
public class SxssfWriteCellHelper extends SimpleWriteCellHelper {

	@Override
	public CellStyle cloneCellStyle(CellStyle srcCellStyle, Workbook desWorkbook) throws Exception {
		if (srcCellStyle instanceof XSSFCellStyle && desWorkbook instanceof XSSFWorkbook) {
			XSSFCellStyle desCellStyle = (XSSFCellStyle) desWorkbook.createCellStyle();
			XSSFCellStyle srcCellStyle_tmp = (XSSFCellStyle) srcCellStyle;
			// 设置背景颜色
			if (srcCellStyle_tmp.getFillBackgroundXSSFColor() != null) {
				desCellStyle.setFillBackgroundColor(srcCellStyle_tmp.getFillBackgroundXSSFColor());
			}
			if (srcCellStyle_tmp.getFillForegroundXSSFColor() != null) {
				desCellStyle.setFillForegroundColor(srcCellStyle_tmp.getFillForegroundXSSFColor());
			}
			desCellStyle.setFillPattern(srcCellStyle_tmp.getFillPattern());
			// 拷贝其他通用属性
			cloneCellStyle(srcCellStyle_tmp, desCellStyle, desWorkbook);

			return desCellStyle;
		} else if (srcCellStyle instanceof XSSFCellStyle && desWorkbook instanceof SXSSFWorkbook) {
			CellStyle desCellStyle = desWorkbook.createCellStyle();
			XSSFCellStyle srcCellStyle_tmp = (XSSFCellStyle) srcCellStyle;
			// 拷贝其他通用属性
			cloneCellStyle(srcCellStyle_tmp, desCellStyle, desWorkbook);
			return desCellStyle;
		}
		return null;
	}


	private void cloneCellStyle(XSSFCellStyle srcCellStyle, CellStyle desCellStyle, Workbook desWorkbook)
			throws Exception {
		if (srcCellStyle != null && desCellStyle != null && desWorkbook != null) {
			// 设置字体
			XSSFFont font = srcCellStyle.getFont();
			Font newFont = desWorkbook.createFont();
			newFont.setFontHeightInPoints(font.getFontHeightInPoints());
			newFont.setFontName(font.getFontName());
			newFont.setBoldweight(font.getBoldweight());
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

}
