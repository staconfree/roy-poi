package com.staconfree.poi.write.model;

import java.util.List;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月4日
 * @since 1.0
 */
public class DataLoopModel {

	public class Mergedregion {
		/**
		 * 这几列的值完全相同的时候才去合并mergeIndex
		 */
		private Integer[] mergeBaseOn;

		/**
		 * 合并的列
		 */
		private Integer[] mergeIndex;

		/**
		 * 上一行的值:{from:"value1",...,to:"value2"}
		 */
		private String lastRowValue;
		/**
		 * 当前行的值
		 */
		private String thisRowValue;

		public Integer[] getMergeBaseOn() {
			return mergeBaseOn;
		}

		public void setMergeBaseOn(Integer[] mergeBaseOn) {
			this.mergeBaseOn = mergeBaseOn;
		}

		public Integer[] getMergeIndex() {
			return mergeIndex;
		}

		public void setMergeIndex(Integer[] mergeIndex) {
			this.mergeIndex = mergeIndex;
		}

		public String getLastRowValue() {
			return lastRowValue;
		}

		public void setLastRowValue(String lastRowValue) {
			this.lastRowValue = lastRowValue;
		}

		public String getThisRowValue() {
			return thisRowValue;
		}

		public void setThisRowValue(String thisRowValue) {
			this.thisRowValue = thisRowValue;
		}

	}

	/**
	 * 从第几行开始写入
	 */
	private int bgnRow;
	/**
	 * 从模板第几行开始
	 */
	private int templateBgnRow;
	/**
	 * 已经循环的行数
	 */
	private int loopedNum = 0;
	/**
	 * 循环的结果集名称
	 */
	private String name;
	/**
	 * 循环体的行(几行几列)，主要是定义样式，模板上支持多行循环
	 */
	private List<List<TemplateCell>> rows;
	/**
	 * 合并单元格
	 */
	private List<Mergedregion> mergedregions;
	
	/**
	 * 行高
	 */
	private short rowHight;

	public int getBgnRow() {
		return bgnRow;
	}

	public void setBgnRow(int bgnRow) {
		this.bgnRow = bgnRow;
	}

	public int getTemplateBgnRow() {
		return templateBgnRow;
	}

	public void setTemplateBgnRow(int templateBgnRow) {
		this.templateBgnRow = templateBgnRow;
	}

	public int getLoopedNum() {
		return loopedNum;
	}

	public void setLoopedNum(int loopedNum) {
		this.loopedNum = loopedNum;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public List<List<TemplateCell>> getRows() {
		return rows;
	}

	public void setRows(List<List<TemplateCell>> rows) {
		this.rows = rows;
	}

	public List<Mergedregion> getMergedregions() {
		return mergedregions;
	}

	public void setMergedregions(List<Mergedregion> mergedregions) {
		this.mergedregions = mergedregions;
	}

	public short getRowHight() {
		return rowHight;
	}

	public void setRowHight(short rowHight) {
		this.rowHight = rowHight;
	}

}
