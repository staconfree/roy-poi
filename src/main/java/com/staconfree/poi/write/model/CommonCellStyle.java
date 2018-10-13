package com.staconfree.poi.write.model;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 单元格通用样式
 * 
 * @author xiexq
 * @date 2014年3月22日
 * @since 1.0
 */
public class CommonCellStyle {
	public static final short FONTHEIGHTINPOINTS = 10;
	public static final String FONTNAME = "宋体";
	public static final short FONTBOLDWEIGHT = 400;
	
	// 设置字体
	private short fontHeightInPoints = FONTHEIGHTINPOINTS;
	private String fontName = FONTNAME;
	private short fontBoldweight = FONTBOLDWEIGHT;
	// 设置边框线型
	// 注：这里不能设置边框颜色，否则会导致导出的单元格不能设置单元格格式
	private short borderTop = 1;
	private short borderBottom = 1;
	private short borderLeft = 1;
	private short borderRight = 1;
	// 设置内容位置:例水平居中,居右，居工
	private short alignment = CellStyle.ALIGN_CENTER;
	// 设置内容位置:例垂直居中,居上，居下
	private short verticalAlignment = CellStyle.VERTICAL_CENTER;
	// 自动换行
	private boolean wrapText = false;
	// 其他属性
	private short dataFormat = 0;
	private boolean hidden = false;
	private short indention = 0;// 首行缩进
	private boolean locked = true;
	private short rotation = 0;// 旋转

	public short getFontHeightInPoints() {
		return fontHeightInPoints;
	}

	public void setFontHeightInPoints(short fontHeightInPoints) {
		this.fontHeightInPoints = fontHeightInPoints;
	}

	public String getFontName() {
		return fontName;
	}

	public void setFontName(String fontName) {
		this.fontName = fontName;
	}

	public short getFontBoldweight() {
		return fontBoldweight;
	}

	public void setFontBoldweight(short fontBoldweight) {
		this.fontBoldweight = fontBoldweight;
	}

	public short getBorderTop() {
		return borderTop;
	}

	public void setBorderTop(short borderTop) {
		this.borderTop = borderTop;
	}

	public short getBorderBottom() {
		return borderBottom;
	}

	public void setBorderBottom(short borderBottom) {
		this.borderBottom = borderBottom;
	}

	public short getBorderLeft() {
		return borderLeft;
	}

	public void setBorderLeft(short borderLeft) {
		this.borderLeft = borderLeft;
	}

	public short getBorderRight() {
		return borderRight;
	}

	public void setBorderRight(short borderRight) {
		this.borderRight = borderRight;
	}

	public short getVerticalAlignment() {
		return verticalAlignment;
	}

	public void setVerticalAlignment(short verticalAlignment) {
		this.verticalAlignment = verticalAlignment;
	}

	public boolean isWrapText() {
		return wrapText;
	}

	public void setWrapText(boolean wrapText) {
		this.wrapText = wrapText;
	}

	public short getDataFormat() {
		return dataFormat;
	}

	public void setDataFormat(short dataFormat) {
		this.dataFormat = dataFormat;
	}

	public boolean isHidden() {
		return hidden;
	}

	public void setHidden(boolean hidden) {
		this.hidden = hidden;
	}

	public short getIndention() {
		return indention;
	}

	public void setIndention(short indention) {
		this.indention = indention;
	}

	public boolean isLocked() {
		return locked;
	}

	public void setLocked(boolean locked) {
		this.locked = locked;
	}

	public short getRotation() {
		return rotation;
	}

	public void setRotation(short rotation) {
		this.rotation = rotation;
	}

	public short getAlignment() {
		return alignment;
	}

	public void setAlignment(short alignment) {
		this.alignment = alignment;
	}

}
