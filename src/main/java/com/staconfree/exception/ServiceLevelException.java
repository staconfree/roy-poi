/**
 * <p>Title: ServiceLevelException.java </p>
 * <p>Description:  </p>
 * <p>Copyright: OrderCenterV1.0 2007 </p>
 * <p>Company: 深圳市泛易网络有限公司 </p>
 */


/**
 * 
 */
package com.staconfree.exception;

import com.staconfree.util.JSONUtil;

import java.util.HashMap;
import java.util.Map;

public class ServiceLevelException extends RuntimeException {

	private static final long serialVersionUID = 988592186626034036L;
	
	private static String getJsonResult(int result, String msg) {
		Map<String, Object> jsonMap = new HashMap<String, Object>();
		jsonMap.put("result", result);
		jsonMap.put("msg", msg);
		return JSONUtil.conversionToJSON(jsonMap);
	}
	
	public ServiceLevelException() {
		super();
	}

	public ServiceLevelException(String arg0) {
		super(arg0);
	}

	public ServiceLevelException(int result, String arg0) {
		super(getJsonResult(result, arg0));
	}
	
	public ServiceLevelException(Throwable arg1) {
		super(arg1);
	}
	
	public ServiceLevelException(String arg0, Throwable arg1) {
		super(arg0,arg1);
	}
}
