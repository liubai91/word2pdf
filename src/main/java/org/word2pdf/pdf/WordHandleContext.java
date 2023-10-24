package org.word2pdf.pdf;

import java.io.InputStream;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.LayoutCollector;
import com.aspose.words.LayoutEnumerator;
import com.lianmed.pdf.TextBoxtHeightListener;

public class WordHandleContext {
	
	private Document doc = null;
	private Object data = null;
	private DocumentBuilder docBuilder = null;
	private LayoutCollector collector = null;
	private LayoutEnumerator enumerator = null;
	private TextBoxtHeightListener listener;
	
	
	
	public WordHandleContext(String fileName, Object data) throws Exception {
		this.doc = new Document(fileName);
		this.data = data;
		this.docBuilder = new DocumentBuilder(doc);
		this.collector = new LayoutCollector(doc);
		this.enumerator = new LayoutEnumerator(doc);
	}
	
	public WordHandleContext(InputStream stream, Object data) throws Exception {
		this.doc = new Document(stream);
		this.data = data;
		this.docBuilder = new DocumentBuilder(doc);
		this.collector = new LayoutCollector(doc);
		this.enumerator = new LayoutEnumerator(doc);
	}
	
	
	public Document getDoc() {
		return doc;
	}
	public void setDoc(Document doc) {
		this.doc = doc;
	}
	public Object getData() {
		return data;
	}
	public void setData(Object data) {
		this.data = data;
	}
	public DocumentBuilder getDocBuilder() {
		return docBuilder;
	}
	public void setDocBuilder(DocumentBuilder docBuilder) {
		this.docBuilder = docBuilder;
	}
	public LayoutCollector getCollector() {
		return collector;
	}
	public void setCollector(LayoutCollector collector) {
		this.collector = collector;
	}
	public LayoutEnumerator getEnumerator() {
		return enumerator;
	}
	public void setEnumerator(LayoutEnumerator enumerator) {
		this.enumerator = enumerator;
	}
	public TextBoxtHeightListener getListener() {
		return listener;
	}
	public void setListener(TextBoxtHeightListener listener) {
		this.listener = listener;
	}
	
	

}
