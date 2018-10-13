package com.staconfree.clazz.reflect;

import com.esotericsoftware.reflectasm.MethodAccess;
import com.staconfree.clazz.ClassUtil;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

/**   
 *
 * @author xiexq
 * @date 2014年4月2日
 * @since 1.0
 */
public class MethodAccessUtil {
	/**
	 * （class及class属性返回的类）的属性工厂类
	 */
	public static Map<String, MethodAccess> _classMethodAccessFactory =new HashMap<String, MethodAccess>();;
	/**
	 * 得到class的fieldAccess
	 * 
	 * @param clazz
	 * @return
	 */
	public static MethodAccess getMethodAccess(Class clazz)  throws Exception{
		if (clazz != null) {
			MethodAccess methodAccess =_classMethodAccessFactory.get(clazz.getName());
			if (null == methodAccess) {
				fillMethodAccess(clazz, 0);
				methodAccess =_classMethodAccessFactory.get(clazz.getName());
			}
			return methodAccess;
		}
		return null;
	}
	

	private static void fillMethodAccess(Class clazz, int level) throws Exception {
//		StringBuilder tipStringBuilder = new StringBuilder();
//		for (int i = 0; i < level; i++) {
//			tipStringBuilder.append(" ");
//		}
//		if (tipStringBuilder.length() > 0) {
//			tipStringBuilder.append(level).append("-");
//		}
//		tipStringBuilder.append(clazz.getName());
//		System.out.println(tipStringBuilder);
		if (!_classMethodAccessFactory.containsKey(clazz.getName())) {
			MethodAccess access = MethodAccess.get(clazz);
			_classMethodAccessFactory.put(clazz.getName(), access);
		} else {
			return;
		}
		Field[] fields = clazz.getDeclaredFields();
		for (Field field : fields) {
			Class fieldType = field.getType();
			if (!ClassUtil.isBasicDataType(fieldType) && !ClassUtil.isCollectionType(fieldType)
					&& !ClassUtil.isArrayType(fieldType) && !ClassUtil.isMapType(fieldType)) {
				// 如果不是(基础数据类型、集合、数组、Map)，递归
				fillMethodAccess(fieldType, level + 1);
			}
		}
	}
}
