package com.staconfree.util;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;


import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;


@SuppressWarnings("unchecked")
public class JSONUtil {
	public static List conversionToList(String json, Class clazz){
		JSONArray jsonArray = JSONArray.fromObject(json);
		List list = (List)JSONArray.toCollection(jsonArray, clazz);
		return list;
	}
	
	public static Object conversionToObject(String json, Class clazz){
		JSONObject jsonObject = JSONObject.fromObject(json);
		return JSONObject.toBean(jsonObject,clazz);
	}
	
	public static String conversionToJSON(Object object) {
		JSONObject jsonObject = JSONObject.fromObject(object);
		return jsonObject.toString();
	}
	
	public static Map conversionToMap(String json){
        JSONObject jsonObject = JSONObject.fromObject(json);  
        Iterator keyIter = jsonObject.keys();
        Map valueMap = new HashMap();

        while( keyIter.hasNext()){
	    	String key = (String)keyIter.next();
	    	Object value = jsonObject.get(key);
	        valueMap.put(key, value);
        }
        
        return valueMap;
    }
	

}

