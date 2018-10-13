package com.staconfree.clazz.model;

/**
 * 
 * 
 * @author xiexq
 * @date 2014年3月3日
 * @since 1.0
 */
public class Student {
	public Integer id;
	public String name;
	public String content3;
	public Teacher teacher;
	
	public Integer getId() {
		return id;
	}
	
	public void setId(Integer id) {
		this.id = id;
	}
	
	public String getName() {
		return name;
	}
	
	public void setName(String name) {
		this.name = name;
	}
	
	public Teacher getTeacher() {
		return teacher;
	}
	
	public void setTeacher(Teacher teacher) {
		this.teacher = teacher;
	}

	public String getContent3() {
		return content3;
	}

	public void setContent3(String content3) {
		this.content3 = content3;
	}
}
