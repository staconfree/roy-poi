package com.staconfree.poi;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月6日
 * @since 1.0
 */
public class PoiConstants {
	public static final String BEGIONLOOP = "$$BEGINLOOP";
	public static final String ENDLOOP = "$$ENDLOOP";
	public static final String LOOP_VARNAME = "name";
	// 是否合并单元格根据哪些列相同
	public static final String LOOP_MERGEDBASEON = "mergedbaseon";
	// mergedbaseon相同的时候才分别合并mergedindex这几列，如果mergedindex为空，则合并mergedbaseon这几列
	public static final String LOOP_MERGEDINDEX = "mergedindex";
	// 数据集的序号,支持加减乘除运算，例如：${#index},${#index + 1},${#index * 2}
	public static final String LOOP_INDEX = "#index";
}
