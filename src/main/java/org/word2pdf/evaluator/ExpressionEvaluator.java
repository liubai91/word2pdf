package org.word2pdf.evaluator;


import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.io.input.ReaderInputStream;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;

import ognl.Ognl;
import ognl.OgnlException;

public class ExpressionEvaluator {
	
	private static Logger logger = LoggerFactory.getLogger(ExpressionEvaluator.class);
	
	private static JSONArray dictionaries = null;
	
	private static Pattern FILTER_PATTERN = Pattern.compile("\\[(.*?)\\]");
	private static Pattern INT_PATTERN = Pattern.compile("\\d*");
	
	static {
		//String filename = ExpressionEvaluator.class.getClassLoader().getResource("dictionaries.json").getFile();

		InputStream is = ExpressionEvaluator.class.getClassLoader().getResourceAsStream("dictionaries.json");
		byte[] bytes = null;
		/*File file = new File(filename);
		String content = "";*/
		try {
			bytes = new byte[is.available()];
			//content = FileUtils.readFileToString(file,"utf-8");
			IOUtils.read(is, bytes);
			dictionaries = JSONArray.parseArray(new String(bytes,"utf-8") );
		} catch (Exception e) {
		}

	}
	
	public static Object evaluateExpression(String expression, Object data) {
		return evaluateExpression(expression, data , -1);
	}
	
	public static Object evaluateExpression(String expression, Object data, int offset) {
		String oriExpression = expression;
		Object result = null;
		expression = expression.trim();
		if(offset>=0) {
			expression = expression.replaceAll("idx", Integer.valueOf(offset).toString());
			//expression.replaceFirst("idx", Integer.valueOf(offset).toString());
		}
		expression = handleFilterGrammar(expression);
		try {
			//expression = expression.replaceAll("\\[", ".toArray()\\[");
			if(expression.contains("->")) {
				if(expression.contains("+")) {
					String[] items = expression.split("\\+");
					items = removeBlankItems(items);
					String ret = "";
					for(String item : items) {
						if(item.contains("->")) {
							ret += handleMap(item.trim(), data);
						}else {
							ret += Ognl.getValue(item.trim(), data);
						}
					}
					result = ret;
				}else {
					result = handleMap(expression, data);
				}
			}else {
				result = Ognl.getValue(expression, data);
			}
		} catch (Exception e) {
			logger.error("表达式解析异常:"+oriExpression);
			logger.error(expression);
			logger.error("------------------");
		}
		result = result==null?"":result;
		return result;
	}
	
	private static String handleMap(String expression,Object data) throws OgnlException {
		String[] dirs = expression.split("->");
		dirs = removeBlankItems(dirs);
		if(dirs.length!=2) {
			return "";
		}
		Object value = Ognl.getValue(dirs[0], data);
		if(value==null) {
			return "";
		}
		//int v = Integer.valueOf(value.toString());
		JSONObject item = findDictionaryItem(dirs[1].trim());
		if(item == null) {
			return "";
		}
		JSONArray enums = item.getJSONArray("enumerations");
		if(enums==null) {
			return "";
		}
		for(int i=0;i<enums.size();i++) {
			JSONObject map = enums.getJSONObject(i);
			if(map.get("value")!=null&&value.toString().trim().equals(map.get("value").toString().trim())) {
				return Objects.toString(map.getString("label"), "");
			}
			/*if(map.getInteger("value")!=null&&v==map.getInteger("value")) {
				return Objects.toString(map.getString("label"), "");
			}*/
		}
		return Objects.toString(value, "");
	}
	
	private static String handleFilterGrammar(String expression) {
		Matcher matcher = FILTER_PATTERN.matcher(expression);
		StringBuffer ret = new StringBuffer();
		while(matcher.find()) {
			String condition = matcher.group(1);
			
			if(!INT_PATTERN.matcher(condition).matches()){
				if(!condition.contains("==")&&condition.contains("=")) {
					condition = condition.replaceAll("=", "==");
				}
				if(matcher.end(0)<expression.length()&&expression.charAt(matcher.end(0))=='.'){
					matcher.appendReplacement(ret, ".{?" + condition + "}[0]");
				}else{
					matcher.appendReplacement(ret, ".{?" + condition + "}");
				}
			}else {
				matcher.appendReplacement(ret, ".toArray()"+matcher.group());
			}
		}
		matcher.appendTail(ret);
		return ret.toString();
	}
	
	

	public static int getSizeOfIterationData(String expression, Object data) {
		int offset = expression.indexOf("idx");
		if(offset == -1) {
			return -1;
		}
		expression = expression.substring(0, offset-1);
		try {
			/*expression = expression.replaceAll("\\[", ".toArray()\\[");
			Object val = Ognl.getValue(expression, data);*/
			Object val =evaluateExpression(expression, data);
			if(val instanceof Set){
				val = ((Set)val).toArray();
			}
			if(val instanceof List) {
				val = ((List) val).toArray();
			}
			if( !val.getClass().isArray() ) {
				return -1;
			}
			Object[] array =  (Object[]) val;
			return array.length;
		} catch (Exception e) {
			logger.error("表达式错误"+expression);
			logger.error(offset+"");
			logger.error(expression);
		}
		return -1;
	}
	
	private static String[] removeBlankItems(String[] items) {
		List<String> list = new ArrayList(Arrays.asList(items));
		Iterator<String> iterator = list.iterator();
		while(iterator.hasNext()) {
			String item = iterator.next();
			if(item.trim().length()==0) {
				iterator.remove();
			}
		}
		items = list.toArray(new String[list.size()]);
		return items;
	}
	
	private static JSONObject findDictionaryItem(String key) {
		if(dictionaries==null) {
			return null;
		}
		String[] ids = key.split("\\.");
		if(ids.length!=2) {
			return null;
		}
		for(int i=0; i<dictionaries.size(); i++) {
			JSONObject item = dictionaries.getJSONObject(i);
			if(item.getString("module").equals(ids[0])&&item.getString("key").equals(ids[1])) {
				return item;
			}
		}
		return null;
	}

	public static JSONArray getDictionaries() {
		return dictionaries;
	}

	public static void setDictionaries(JSONArray dictionaries) {
		ExpressionEvaluator.dictionaries = dictionaries;
	}

}