package com.staconfree.var;

import com.staconfree.clazz.reflect.DataReflect;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 变量转换器
 * 
 * @author xiexq
 * @date 2014年3月5日
 * @since 1.0
 */
public class VarConverter {
	private static final String LOOPINDEX = "#index";

	/**
	 * 把循环数据data里面包含的变量转换为数据
	 * 
	 * @param var
	 *            变量，如："${id}${name}"
	 * @param data
	 * @param dataReflect
	 *            data对应的class的类反射，所有的data类型是一样的，作为参数是为了只初始化一次
	 * @param dataStack
	 * @param dataIndex
	 *            本次数据的索引
	 * @param pastDataNum
	 *            以前数据的总数
	 * @return
	 */
	public static String convertLoopData(String var, Object data, DataReflect dataReflect,
			Map<String, Object> dataStack, int dataIndex, int pastDataNum) throws Exception {
		List<String> attributes = VarAnalysiser.getVarNames(var);
		if (attributes != null && !attributes.isEmpty() && data != null && dataReflect != null) {
			Map<String, String> attributeValues = new HashMap<String, String>();
			for (String attributeName : attributes) {
				attributeName = attributeName.trim();
				if (attributeName.startsWith(LOOPINDEX)) {// 如果是取索引
					int thisIndex = calcDataIndex(dataIndex + pastDataNum,
							attributeName.substring(attributeName.indexOf(LOOPINDEX) + LOOPINDEX.length()));
					attributeValues.put(attributeName, String.valueOf(thisIndex));
					continue;
				}
				// 根据变量名取值
				Object changeValue = dataReflect.invokeGetMethod(attributeName, data);
				if (changeValue != null) {
					attributeValues.put(attributeName, changeValue.toString());
				} else {
					if (dataStack != null) {
						// 到数据栈里面找
						changeValue = convertAttribute(attributeName, dataStack);
						if (changeValue != null) {
							attributeValues.put(attributeName, changeValue.toString());
						} else {
							attributeValues.put(attributeName, "");
						}
					} else {
						attributeValues.put(attributeName, "");
					}
				}
			}
			String result = VarAnalysiser.getVarWithValues(attributeValues, var);
			return result;
		}
		return var;
	}

	private static int calcDataIndex(int index, String calcValue) throws Exception {
		String calc = "";
		int value = 0;
		try {
			if (calcValue.trim().startsWith("+")) {
				calc = "+";
				value = Integer.valueOf(calcValue.substring(calcValue.indexOf("+") + 1).trim());
			} else if (calcValue.trim().startsWith("-")) {
				calc = "-";
				value = Integer.valueOf(calcValue.substring(calcValue.indexOf("-") + 1).trim());
			} else if (calcValue.trim().startsWith("*")) {
				calc = "*";
				value = Integer.valueOf(calcValue.substring(calcValue.indexOf("*") + 1).trim());
			} else if (calcValue.trim().startsWith("/")) {
				calc = "/";
				value = Integer.valueOf(calcValue.substring(calcValue.indexOf("/") + 1).trim());
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return calcDataIndex(index, calc, value);
	}

	private static int calcDataIndex(int index, String calc, int value) throws Exception {
		if ("+".equals(calc)) {
			return index + value;
		} else if ("-".equals(calc)) {
			return index - value;
		} else if ("*".equals(calc)) {
			return index * value;
		} else if ("/".equals(calc)) {
			return index / value;
		}
		return index;
	}

	private static Object convertAttribute(String attributeName, Map<String, Object> dataStack) throws Exception {
		if (attributeName != null && dataStack != null) {
			Object obj = dataStack.get(attributeName);
			if (obj != null) {
				return obj;
			} else {
				if (attributeName.contains(".")) {
					String parent = attributeName.substring(0, attributeName.indexOf("."));
					String child = attributeName.substring(attributeName.indexOf(".") + 1);
					if (parent != null && !"".equals(parent) && child != null && !"".equals(child)) {
						Object parentObj = dataStack.get(parent);
						if (parentObj != null) {
							Class clazz = parentObj.getClass();
							// 根据变量名取值
							DataReflect dataReflect = new DataReflect(clazz);
							Object childObj = dataReflect.invokeGetMethod(child, parentObj);
							if (childObj != null) {
								return childObj;
							}
						}
					}
				}
			}
		}
		return null;
	}

	/**
	 * 普通变量的转换
	 * 
	 * @param var
	 *            变量，如："${id}${name}"
	 * @param dataStack
	 * @return
	 */
	public static String convertCommonData(String var, Map<String, Object> dataStack) throws Exception {
		List<String> attributes = VarAnalysiser.getVarNames(var);
		if (dataStack == null) {
			return var;
		}
		if (attributes != null && !attributes.isEmpty()) {
			Map<String, String> attributeValues = new HashMap<String, String>();
			for (String attributeName : attributes) {
				Object changeValue = convertAttribute(attributeName, dataStack);
				if (changeValue != null) {
					attributeValues.put(attributeName, changeValue.toString());
				} else {
					attributeValues.put(attributeName, "");
				}
			}
			String result = VarAnalysiser.getVarWithValues(attributeValues, var);
			return result;
		}
		return var;
	}

	public static void main(String[] args) {
		try {
			System.out.println(VarConverter.convertSystemData("23"));
			String str = "-$$yyyy-$$MM-$$dd-";
			System.out.println(VarConverter.replaceSystemData(str.toCharArray(), new Date()));
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 系统变量的转换
	 * 
	 * @param var
	 *            变量，如："$$yyyy-$$MM-$$dd $$hh:$$mm:$$ss"
	 * @return
	 */
	public static String convertSystemData(String var) throws Exception {
		if (var != null && !"".equals(var)) {
			Date date = new Date();
			char[] charArr = var.toCharArray();
			return replaceSystemData(charArr, date);
		}
		return var;
	}

	private static String replaceSystemData(char[] charArr, Date date) {
		if (charArr != null && charArr.length > 0 && date != null) {
			int length = charArr.length;
			String result = "";
			boolean isReplace = false;
			for (int i = 0; i < length; i++) {
				if (charArr[i] == '$' && i < length - 5 && charArr[i + 1] == '$' && charArr[i + 2] == 'y'
						&& charArr[i + 3] == 'y' && charArr[i + 4] == 'y' && charArr[i + 5] == 'y') {// 年
					isReplace = true;
					if (i != 0) {
						char[] beforeArr = new char[i];
						System.arraycopy(charArr, 0, beforeArr, 0, i);
						result = result + String.valueOf(beforeArr);
					}
					String replaceStr = new SimpleDateFormat("yyyy").format(date);
					result = result + replaceStr;
					int afterSize = length - 1 - (i + 5);
					if (afterSize != 0) {
						char[] afterArr = new char[afterSize];
						System.arraycopy(charArr, i + 5 + 1, afterArr, 0, afterSize);
						return result + replaceSystemData(afterArr, date);
					} else {
						return result;
					}
				} else if (charArr[i] == '$' && i < length - 3) {
					DateFormat dateFormat = null;
					if (charArr[i + 1] == '$' && charArr[i + 2] == 'M' && charArr[i + 3] == 'M') {// 月
						dateFormat = new SimpleDateFormat("MM");
						isReplace = true;
					} else if (charArr[i + 1] == '$' && charArr[i + 2] == 'd' && charArr[i + 3] == 'd') {// 日
						dateFormat = new SimpleDateFormat("dd");
						isReplace = true;
					} else if (charArr[i + 1] == '$' && charArr[i + 2] == 'h' && charArr[i + 3] == 'h') {// 时
						dateFormat = new SimpleDateFormat("hh");
						isReplace = true;
					} else if (charArr[i + 1] == '$' && charArr[i + 2] == 'H' && charArr[i + 3] == 'H') {// 时
						dateFormat = new SimpleDateFormat("HH");
						isReplace = true;
					} else if (charArr[i + 1] == '$' && charArr[i + 2] == 'm' && charArr[i + 3] == 'm') {// 分
						dateFormat = new SimpleDateFormat("mm");
						isReplace = true;
					} else if (charArr[i + 1] == '$' && charArr[i + 2] == 's' && charArr[i + 3] == 's') {// 秒
						dateFormat = new SimpleDateFormat("ss");
						isReplace = true;
					}
					if (dateFormat != null) {
						if (i != 0) {
							char[] beforeArr = new char[i];
							System.arraycopy(charArr, 0, beforeArr, 0, i);
							result = result + String.valueOf(beforeArr);
						}
						String replaceStr = dateFormat.format(date);
						result = result + replaceStr;
						int afterSize = length - 1 - (i + 3);
						if (afterSize != 0) {
							char[] afterArr = new char[afterSize];
							System.arraycopy(charArr, i + 3 + 1, afterArr, 0, afterSize);
							return result + replaceSystemData(afterArr, date);
						} else {
							return result;
						}
					}
					i = i + 3;
				}
			}
			if (!isReplace) {
				result = String.valueOf(charArr);
				return result;
			}
		}
		return "";
	}

}
