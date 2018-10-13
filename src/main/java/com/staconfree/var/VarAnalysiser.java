package com.staconfree.var;

import java.util.*;

/**
 * 变量解析器
 * 
 * @author xiexq
 * @date 2014年3月5日
 * @since 1.0
 */
public class VarAnalysiser {
	public static void main(String[] args)  throws Exception{
		String name = "${name}${name}";
		System.out.println(VarAnalysiser.getVarNames(name));
		Map<String, String> attributeValues = new HashMap<String, String>();
		attributeValues.put("id", "123123");
		System.out.println("${id}${id}".replaceAll("\\$\\{id\\}", "123123"));
		System.out.println("${#index * 1}".replaceAll("\\$\\{#index * 1\\}", "123123"));
		System.out.println(VarAnalysiser.getVarWithValues(attributeValues, "${id}"));
	}
	
	/**
	 * 得到变量里面的 "${name}" 此类的
	 * 
	 * @param formula
	 * @return
	 */
	public static List<String> getVarNames(String formula)  throws Exception{
		if (formula != null) {
			List<String> list = new ArrayList<String>();
			String restFormula = formula;
			int i = restFormula.indexOf("${");
			int j = restFormula.indexOf("}");
			while (i >= 0 && j > 0 && i < j) {
				String attribute = restFormula.substring(i + 2, j);
				if (!list.contains(attribute)) {
					list.add(attribute);
				}
				restFormula = restFormula.substring(j + 1);
				i = restFormula.indexOf("${");
				j = restFormula.indexOf("}");
			}
			return list;
		}
		return null;
	}
	
	/**
	 * 替换变量里面的值
	 * 
	 * @param attributeValues
	 * @param formula
	 * @return
	 */
	public static String getVarWithValues(Map<String, String> attributeValues, String formula) {
		if (attributeValues != null && !attributeValues.isEmpty()) {
			Set<String> keySet = attributeValues.keySet();
			for (String key : keySet) {
				// 先屏蔽+号和*号两个特殊字符
				formula = formula.replaceAll("\\*", "").replaceAll("\\+", "");
				String replaceAttribute = key.replaceAll("\\*", "").replaceAll("\\+", "");
				formula = formula.replaceAll("\\$\\{" + replaceAttribute + "\\}", attributeValues.get(key));
			}
			return formula;
		}
		return null;
	}
	
}
