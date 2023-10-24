package org.word2pdf.processor;

import com.lianmed.pdf.WordHandleContext;

public interface Processor {
	
	public void process(WordHandleContext context) throws Exception;

}
