package org.word2pdf.core;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Matcher;

import com.aspose.words.BookmarkEnd;
import com.aspose.words.BookmarkStart;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.LayoutCollector;
import com.aspose.words.LayoutEnumerator;
import com.aspose.words.Paragraph;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.Run;
import com.aspose.words.Shape;
import com.lianmed.evaluator.ExpressionEvaluator;
import com.lianmed.internal.PlaceHolder;
import com.lianmed.pdf.WordHandleContext;
import com.lianmed.text.TextAlign;
import com.lianmed.util.NodeUtil;

public class TextBoxCoverReplace {
	
	public void method(WordHandleContext context) throws Exception {
		
		Map<PlaceHolder,Matcher> runs = null;
		
		Document doc = context.getDoc();
		DocumentBuilder builder = context.getDocBuilder();
		Object data = context.getData();
		LayoutCollector collector = context.getCollector();
		LayoutEnumerator enumerator = context.getEnumerator();
		
		Map<BookmarkStart,Matcher> bookmarks = new LinkedHashMap<BookmarkStart,Matcher>();
		Map<BookmarkStart,BookmarkEnd> bookmarkpairs = new HashMap<BookmarkStart,BookmarkEnd>();
		Map<BookmarkStart,Run> bookmarkToRunMapper = new HashMap<BookmarkStart,Run>();
		for(PlaceHolder item : runs.keySet()) {
			Run run = item.getStartNode();
			Matcher matcher = runs.get(item);
			String format = matcher.group(1).trim();
			String expression = matcher.group(2).trim();
			String value = Objects.toString(ExpressionEvaluator.evaluateExpression(expression, data), "");
			//String value = SimpleValueEvaluator.evaluateEL(expression, data);
			if(format.equalsIgnoreCase("L")||format.equalsIgnoreCase("")||format.equalsIgnoreCase("M")||format.equalsIgnoreCase("R")) {
				
				String bmName = String.format("bm_%s_%d", expression , run.hashCode());
				builder.moveTo(run);
				BookmarkStart start = builder.startBookmark(bmName);
				BookmarkEnd end = builder.endBookmark(bmName);
				run.getParentNode().insertAfter(end, item.getEndNode());
				bookmarks.put(start, matcher);
				bookmarkpairs.put(start, end);
				bookmarkToRunMapper.put(start, run);
			}else {
				run.setText(value);
			}
		}
		
		doc.updatePageLayout();
		/*LayoutCollector collector = new LayoutCollector(doc);
		LayoutEnumerator enumerator = new LayoutEnumerator(doc);*/
		
		for(BookmarkStart bmStart : bookmarks.keySet()) {
			Shape textBoxShape = NodeUtil.insertTextBoxCoverBookmark(builder, collector, enumerator, bmStart, bookmarkpairs.get(bmStart));
			Paragraph paragraph = textBoxShape.getLastParagraph();
			paragraph.getParagraphFormat().setAlignment(ParagraphAlignment.LEFT);
			Run run = (Run) bookmarkToRunMapper.get(bmStart).deepClone(true);
			Matcher matcher = bookmarks.get(bmStart);
			String format = matcher.group(1).trim();
			String expression = matcher.group(2).trim();
			String value = Objects.toString(ExpressionEvaluator.evaluateExpression(expression, data), "");
			//String value = SimpleValueEvaluator.evaluateEL(expression, data);
			//value = expression;
			// assist debug
			/*if(value==null||value.trim().length()==0) {
				value = expression;
			}*/
			String origin = run.getText();
			run.setText(TextAlign.formatText(TextAlign.LEFT, origin, value, run.getFont()));
			paragraph.appendChild(run);
		}
	}

}
