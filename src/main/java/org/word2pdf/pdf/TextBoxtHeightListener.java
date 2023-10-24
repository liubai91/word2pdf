package org.word2pdf.pdf;

public interface TextBoxtHeightListener {
	
	/**
	 * 
	 * @param origin 原始高度
	 * @param adaptive 适应高度
	 */
	public void fitShapeToText(double origin,double adaptive) ;

}
