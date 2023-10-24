package org.word2pdf.util;

public class StringUtil {

	public static int search(String str, String strRes) {
		int n = 0;// 计数器
		int index = 0;// 指定字符的长度
		index = str.indexOf(strRes);
		while (index != -1) {
			n++;
			index = str.indexOf(strRes, index + 1 + strRes.length());
		}

		return n;
	}
	
	/**
	 *   获取第n次字串的索引
	 * @param str
	 * @param i
	 * @return
	 */
	public static int indexOf(String str, String strRes, int i) {
		int n = -1;
		int index = 0;
		index = str.indexOf(strRes);
		while (index != -1 &&  n<i) {
			n++;
			index = str.indexOf(strRes, index + 1 + strRes.length());
		}
		if(n == i ) {
			return index;
		}
		return -1;
	}

}
