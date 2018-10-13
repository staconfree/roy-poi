package com.staconfree.clazz.reflect;

import com.esotericsoftware.reflectasm.FieldAccess;
import com.esotericsoftware.reflectasm.MethodAccess;
import com.staconfree.clazz.model.Student;
import com.staconfree.clazz.model.Teacher;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月3日
 * @since 1.0
 */
public class DataReflect {

	private Class<?> _clazz;

	/**
	 * 
	 */
	public DataReflect(Class clazz) {
		this._clazz = clazz;
	}

	public Object getPublicFieldValue(String column, Object obj) throws Exception {
		return getPublicFieldValue(column, obj, _clazz);
	}

	/**
	 * 
	 * @param column
	 *            字段。"id"，代表取obj里面的id字段；"teacher.name"代表取obj里面的teacher对象的name字段
	 * @param obj
	 *            数据
	 * @param clazz
	 *            数据格式，如Student.class
	 */
	private Object getPublicFieldValue(String column, Object obj, Class clazz) throws Exception {
		if (column == null || column.equals("")) {
			return null;
		}
		FieldAccess access = FieldAccessUtil.getFieldAccess(clazz);
		String attribute = null;// 属性字段
		String restColumn = null;// 剩余的字段
		int hasDot = column.indexOf(".");
		if (hasDot < 0) {
			attribute = column;
		} else {
			attribute = column.substring(0, hasDot);
			restColumn = column.substring(hasDot + 1);
		}
		if (attribute != null && !attribute.equals("")) {
			Object value_tmp = null;
			Class clazz_tmp = null;
			try {
				value_tmp = access.get(obj, attribute);
				clazz_tmp = value_tmp.getClass();
			} catch (Exception e) {
				// e.printStackTrace();
			}
			if (restColumn != null && !restColumn.equals("")) {
				return getPublicFieldValue(restColumn, value_tmp, clazz_tmp);
			} else {
				return value_tmp;
			}
		} else {
			return null;
		}
	}

	public Object invokeGetMethod(String column, Object obj) throws Exception {
		return invokeGetMethod(column, obj, _clazz);
	}

	private Object invokeGetMethod(String column, Object obj, Class clazz) throws Exception {
		if (column == null || column.equals("")) {
			return null;
		}
		MethodAccess methodAccess = MethodAccessUtil.getMethodAccess(clazz);
		String attribute = null;// 属性字段
		String restColumn = null;// 剩余的字段
		int hasDot = column.indexOf(".");
		if (hasDot < 0) {
			attribute = column;
		} else {
			attribute = column.substring(0, hasDot);
			restColumn = column.substring(hasDot + 1);
		}
		if (attribute != null && !attribute.equals("")) {
			Object value_tmp = null;
			Class clazz_tmp = null;
			try {
				value_tmp = methodAccess.invoke(obj, changeToGetMethodName(attribute));
				clazz_tmp = value_tmp.getClass();
			} catch (Exception e) {
				// e.printStackTrace();
			}
			if (restColumn != null && !restColumn.equals("") && value_tmp != null) {
				return invokeGetMethod(restColumn, value_tmp, clazz_tmp);
			} else {
				return value_tmp;
			}
		} else {
			return null;
		}
	}

	private String changeToGetMethodName(String str) {
		String firstLeter = str.substring(0, 1);
		String other = str.substring(1);
		return new StringBuilder("get").append(firstLeter.toUpperCase()).append(other).toString();
	}

	public static void main(String[] args) throws Exception {
		Teacher teacher = new Teacher();
		teacher.setId(1);
		teacher.setName("老师1");
		Student student = new Student();
		student.setId(1);
		student.setName("学生1");
		student.setTeacher(teacher);

		DataReflect dataReflect = new DataReflect(Student.class);
		System.out.println("最终结果：" + dataReflect.getPublicFieldValue("teacher.name", student));
		System.out.println("最终结果：" + dataReflect.invokeGetMethod("teacher.name", student));
	}

}
