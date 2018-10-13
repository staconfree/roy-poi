package com.staconfree.clazz.reflect;

import com.esotericsoftware.reflectasm.FieldAccess;
import com.staconfree.clazz.ClassUtil;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月3日
 * @since 1.0
 */
public class FieldAccessUtil {
	/**
	 * （class及class属性返回的类）的属性工厂类
	 */
	public static Map<String, FieldAccess> _classFieldAccessFactory =new HashMap<String, FieldAccess>();;
	
	/**
	 * 得到class的fieldAccess
	 * 
	 * @param clazz
	 * @return
	 */
	public static FieldAccess getFieldAccess(Class clazz)  throws Exception{
		if (clazz != null) {
			FieldAccess fieldAccess =_classFieldAccessFactory.get(clazz.getName());
			if (null == fieldAccess) {
				fillFieldAccess(clazz, 0);
				fieldAccess =_classFieldAccessFactory.get(clazz.getName());
			}
			return fieldAccess;
		}
		return null;
	}
	
	private static void fillFieldAccess(Class clazz, int level)  throws Exception{
//		StringBuilder tipStringBuilder = new StringBuilder();
//		for (int i = 0; i < level; i++) {
//			tipStringBuilder.append(" ");
//		}
//		if (tipStringBuilder.length() > 0) {
//			tipStringBuilder.append(level).append("-");
//		}
//		tipStringBuilder.append(clazz.getName());
//		System.out.println(tipStringBuilder);
		if (!_classFieldAccessFactory.containsKey(clazz.getName())) {
			FieldAccess access = FieldAccess.get(clazz);
			_classFieldAccessFactory.put(clazz.getName(), access);
		} else {
			return;
		}
		Field[] fields = clazz.getDeclaredFields();
		for (Field field : fields) {
			Class fieldType = field.getType();
			if (!ClassUtil.isBasicDataType(fieldType) && !ClassUtil.isCollectionType(fieldType)
					&& !ClassUtil.isArrayType(fieldType) && !ClassUtil.isMapType(fieldType)) {
				// 如果不是(基础数据类型、集合、数组、Map)，递归
				fillFieldAccess(fieldType, level + 1);
			}
		}
	}
	
}
