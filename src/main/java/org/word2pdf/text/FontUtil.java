package org.word2pdf.text;

import java.util.HashMap;
import java.util.Map;


public class FontUtil {
	
	final private static Map<String,String> fontInfo = new HashMap<String,String>();
	
	static {
		fontInfo.put("宋体", "SimSun");
		fontInfo.put("黑体", "SimHei");
		fontInfo.put("微软雅黑", "MicrosoftYaHei");
		fontInfo.put("微软雅黑 Light", "MicrosoftYaHeiLight");
		fontInfo.put("楷体", "KaiTi");
		fontInfo.put("仿宋", "FangSong");
		fontInfo.put("新宋体", "NSimSun");
		fontInfo.put("Calibri Light", "Calibri-Light");
	}
	
	/*public static Font findFont(String fontName) {
		if(fontInfo.containsKey(fontName)) {
			fontName = fontInfo.get(fontName);
		}
		return FontRepository.findFont(fontName);
	}
*/
}