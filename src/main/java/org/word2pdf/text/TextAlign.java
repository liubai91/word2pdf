package org.word2pdf.text;

import java.awt.FontMetrics;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;

import com.aspose.words.Font;

public class TextAlign {
	
	public final static int LEFT = 0;
	public final static int MIDDLE = 1;
	public final static int RIGHT = 2;
	private final static String unitPattern = "[il\\{\\}\\[\\]\\(\\)]";
	private final static String dunitPattern = "^[\u0000-\u007F]{1}$";
	
	public static String formatText(int alignStyle,String orignal,String dest) {
		
		StringBuffer sb = new StringBuffer();
		if(alignStyle==MIDDLE) {
			int total = calculateRelativeWidth(orignal);
			//System.out.println(String.format("原始字符串长度%d，计算得到的单元长度%d", orignal.length(), total));
			int cur = calculateRelativeWidth(dest);
			int gap = total - cur ;
			if(gap%2==0) {
				int num = gap/4;
				if(gap%4==0) {
					sb.append(new String(new char[num]).replace("\0", " "));
					sb.append(dest);
					sb.append(new String(new char[num+1]).replace("\0", " "));
				}else {
					sb.append(new String(new char[num+1]).replace("\0", " "));
					sb.append(dest);
					sb.append(new String(new char[num+1]).replace("\0", " "));
				}
			}
			
		}
		return sb.toString();
	}
	
	//相对于il{}[]()等字符的长度
	public static int calculateRelativeWidth(String text) {
		int len = 0;
		char[] chars = text.toCharArray();
		String a = "@";
		for(char c : chars) {
			 if(c=='@') {
				 len = len + 4;
			 }else if(Character.toString(c).matches(unitPattern)) {
				 len = len + 1;
			 }else if(Character.toString(c).matches(dunitPattern)) {
				 len = len + 2;
			 }else {
				 len = len + 4;
			 }
		}
		return len;
	}

	public static String formatText(int alignStyle, String ori, String des, Font font) {
		StringBuffer sb = new StringBuffer();
		sb.append(des);
		int counter = 0;
		while(counter<100) {
			sb.append(" ");
			counter++;
		}
		/*double owidth = measureString(font,ori);
		if(owidth>measureString(font,sb.toString())) {
			while(owidth>measureString(font,sb.toString())) {
				if(alignStyle == LEFT) {
					sb.append(' ');
				}else if(alignStyle == MIDDLE) {
					if(counter%2==0) {
						sb.insert(0, ' ');
					}else {
						sb.append(' ');
					}
					
				}else if(alignStyle == RIGHT) {
					sb.insert(0, ' ');
				}
				counter++;
			}
		}*//*else if(owidth<measureString(font,sb.toString())) {
			
			do {
				int fontSize = (int) (font.getSize() - 0.5);
				font.setSize(fontSize);
			}while(owidth<measureString(font,sb.toString()));
			
		}*/
		return sb.toString();
	}
	
	private static double measureString(Font font,String str) {
		BufferedImage image = new BufferedImage(1, 1, BufferedImage.TYPE_INT_ARGB);
		Graphics2D g = image.createGraphics();
		
		FontMetrics fm = g.getFontMetrics(new java.awt.Font(font.getName(), java.awt.Font.PLAIN, (int)font.getSize()));
		return fm.stringWidth(str);
	}
}