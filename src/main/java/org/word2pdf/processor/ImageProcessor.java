package org.word2pdf.processor;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.aspose.words.Document;
import com.aspose.words.FindReplaceDirection;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.words.ReportingEngine;
import com.lianmed.evaluator.ExpressionEvaluator;
import com.lianmed.pdf.WordHandleContext;
import com.lianmed.processor.Processor;

public class ImageProcessor implements IReplacingCallback, Processor {
	
	private static final Pattern IMG_DIRECTIVE_PATTERN = Pattern.compile("(<<image.*?\\[)(.*)(\\].*?>>)");
	
	private List<Object> imgSource;
	private List<String> imgSourceName;
	
	private WordHandleContext context;
	
	@Override
	public void process(WordHandleContext context) throws Exception {
		this.context = context;
		imgSource = new ArrayList<Object>();
		imgSourceName = new ArrayList<String>();
		Document doc = context.getDoc();
		FindReplaceOptions options = new FindReplaceOptions();
		options.setDirection(FindReplaceDirection.FORWARD);
		options.setReplacingCallback(this);
		doc.getRange().replace(IMG_DIRECTIVE_PATTERN, "", options);
		
		try {
			if(imgSource.size()>0&&imgSource.size()==imgSourceName.size()) {
				ReportingEngine engine = new ReportingEngine();
				engine.buildReport(doc, imgSource.toArray() , imgSourceName.toArray(new String[0]));
			}else {
				
				ReportingEngine engine = new ReportingEngine();
				engine.buildReport(doc, context.getData());
			}
		} catch (Exception e) {
			
		}
		
	}

	@Override
	public int replacing(ReplacingArgs e) throws Exception {
		Object data = context.getData();
		Matcher matcher = e.getMatch();
		
		Object value = ExpressionEvaluator.evaluateExpression(e.getMatch().group(2), data);
		//Object value = SimpleValueEvaluator.getProperty(e.getMatch().group(2), data , -1);
		if(value!=null && value instanceof InputStream) {
			int idx = imgSource.size();
			imgSource.add(value);
			imgSourceName.add(String.format("img%d", idx));
			String txt = matcher.replaceFirst(String.format("$1img%d$3", idx));
			e.setReplacement(txt);
		}else {
			e.setReplacement("");
		}
		return ReplaceAction.REPLACE;
	}
}
