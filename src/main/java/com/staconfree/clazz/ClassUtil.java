package com.staconfree.clazz;

import java.sql.Date;
import java.sql.Timestamp;
import java.util.Collection;
import java.util.Map;
import java.util.Set;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月5日
 * @since 1.0
 */
public class ClassUtil {
	
	/**
	 * 是否是基础数据类型
	 * 
	 * @param clazz
	 * @return
	 */
	public static boolean isBasicDataType(Class clazz) {
		if (clazz == double.class || clazz == Double.class || clazz == float.class || clazz == Float.class
				|| clazz == String.class || clazz == short.class || clazz == Short.class || clazz == int.class
				|| clazz == Integer.class || clazz == long.class || clazz == Long.class || clazz == boolean.class
				|| clazz == Boolean.class) {
			return true;
		} else if (clazz == Date.class || clazz == Timestamp.class || clazz == java.util.Date.class) {
			return true;
		}
		return false;
	}

	public static boolean isCollectionType(Class clazz) {
		if (Collection.class.isAssignableFrom(clazz)) {
			return true;
		}
		return false;
	}

	public static boolean isArrayType(Class clazz) {
		if (clazz.isArray()) {
			return true;
		}
		return false;
	}
	
	public static boolean isMapType(Class clazz) {
		if (Map.class.isAssignableFrom(clazz)) {
			return true;
		}
		return false;
	}
	
	public static void main(String[] args) {
		System.out.println(ClassUtil.isCollectionType(Set.class));
	}
}
