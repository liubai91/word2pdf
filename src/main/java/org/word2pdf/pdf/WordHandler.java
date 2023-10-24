package org.word2pdf.pdf;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import com.lianmed.pdf.TextBoxtHeightListener;
import com.lianmed.pdf.WordHandleContext;
import com.lianmed.processor.*;
import org.apache.commons.io.FileUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.alibaba.fastjson.JSONArray;
import com.aspose.words.Document;
import com.aspose.words.LoadOptions;
import com.aspose.words.SaveFormat;
import com.lianmed.evaluator.ExpressionEvaluator;
import com.lianmed.lisense.License;

/**
 * $O{}这种格式溢出换行问题待解决
 * @author nhl220
 *
 */
public class WordHandler {
	
	private final Logger log = LoggerFactory.getLogger(WordHandler.class);
	
	private WordHandleContext context = null;
	
	public WordHandler(String wordFile, Object data) throws Exception {
		License.getWordLicense();
		this.context = new WordHandleContext(wordFile, data);
	}
	
	public WordHandler(InputStream is, Object data) throws Exception {
		License.getWordLicense();
		this.context = new WordHandleContext(is, data);
	}
	
	public WordHandler(String wordFile, Object data, String fontpath) throws Exception {
		License.getWordLicense();
		LoadOptions loadOptions = setupFont(fontpath);
		this.context = new WordHandleContext(wordFile, data);
	}
	
	public WordHandler(InputStream is, Object data, String fontpath) throws Exception {
		License.getWordLicense();
		LoadOptions loadOptions = setupFont(fontpath);
		this.context = new WordHandleContext(is, data);
	}

	private LoadOptions setupFont(String fontpath) {
		/*FolderFontSource fs = new FolderFontSource(fontpath);
		FontRepository.getSources().add(fs);
		FontSettings fontSettings = new FontSettings();
		fontSettings.setFontsFolder(fontpath, true);
		LoadOptions loadOptions  = new LoadOptions();
		loadOptions.setFontSettings(fontSettings);*/
		return null;
	}
	
	
	public Document process(String destFile) throws Exception {
		FileOutputStream os = new FileOutputStream(destFile);   
		return process(os);
	}
	
	public Document process(OutputStream os) throws Exception {
		List<Processor> processors = new ArrayList<Processor>();
		//processors.add(new IndentProcessor());

		processors.add(new GroovyProcessor());
		processors.add(new TableProcessor());
		processors.add(new ConditionalBlockProcessor());
		processors.add(new ImageProcessor());
		processors.add(new CommentProcessor());
		processors.add(new DataProcessor());
		processors.add(new BlankPageRemoveProcessor());
		for(Processor processor : processors) {
			processor.process(context);
		}
		
		Document doc = context.getDoc();
		doc.updatePageLayout();
		//doc.save("d:/b/test1.doc");
		if(os!=null) {
			doc.save(os, SaveFormat.PDF);
		}
		return doc;
	}

	public Document process(OutputStream os,int saveFormat) throws Exception {
		List<Processor> processors = new ArrayList<Processor>();
		//processors.add(new IndentProcessor());

		processors.add(new GroovyProcessor());
		processors.add(new TableProcessor());
		processors.add(new ConditionalBlockProcessor());
		processors.add(new ImageProcessor());
		processors.add(new CommentProcessor());
		processors.add(new DataProcessor());
		processors.add(new BlankPageRemoveProcessor());
		for(Processor processor : processors) {
			processor.process(context);
		}

		Document doc = context.getDoc();
		doc.updatePageLayout();
		//doc.save("d:/b/test1.doc");
		if(os!=null) {
			doc.save(os, saveFormat);
		}
		return doc;
	}
	
	public void setDictionary(File file) {
		String content = "";
		try {
			content = FileUtils.readFileToString(file,"utf-8");
		} catch (IOException e) {
		}
		JSONArray dictionaries = JSONArray.parseArray(content);
		ExpressionEvaluator.setDictionaries(dictionaries);
	}

	public TextBoxtHeightListener getListener() {
		return context.getListener();
	}

	public void setListener(TextBoxtHeightListener listener) {
		context.setListener(listener);
	}
	
}
