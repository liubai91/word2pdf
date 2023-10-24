package org.word2pdf.processor;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.aspose.words.BookmarkEnd;
import com.aspose.words.BookmarkStart;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FindReplaceDirection;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.LayoutCollector;
import com.aspose.words.LayoutEnumerator;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.words.Run;
import com.aspose.words.Shape;
import com.lianmed.core.PlaceHolderResolver;
import com.lianmed.evaluator.ExpressionEvaluator;
import com.lianmed.pdf.WordHandleContext;
import com.lianmed.processor.Processor;
import com.lianmed.text.TextAlign;
import com.lianmed.util.NodeUtil;

/**
 * textbox 越界问题待处理
 * @author nhl220
 *
 */
public class DataProcessor implements IReplacingCallback, Processor {
	
	private static final Pattern DATA_ACCESS_DIRECTIVE_PATTERN = Pattern.compile("\\$(.?)\\{(.*?)\\}");
	
	private WordHandleContext context;
	private Map<Run,Matcher> runs;
	
	@Override
	public void process(WordHandleContext context) throws Exception {
		this.context = context;
		runs = new LinkedHashMap<Run,Matcher>();
		replaceExpressionBeginWith$();
	}
	
	@Override
	public int replacing(ReplacingArgs e) throws Exception {
		Run currentNode =  PlaceHolderResolver.resolveRunWithUnderline(e);
		runs.put(currentNode, e.getMatch());
		return ReplaceAction.SKIP;
	}
	
	private void replaceExpressionBeginWith$() throws Exception {
		
		Document doc = context.getDoc();
		DocumentBuilder builder = context.getDocBuilder();
		Object data = context.getData();
		LayoutCollector collector = context.getCollector();
		LayoutEnumerator enumerator = context.getEnumerator();
		
		FindReplaceOptions options = new FindReplaceOptions();
		options.setDirection(FindReplaceDirection.BACKWARD);
		options.setReplacingCallback(this);
		doc.getRange().replace(DATA_ACCESS_DIRECTIVE_PATTERN, "", options);
		
		
		Map<BookmarkStart,Matcher> bookmarks = new LinkedHashMap<BookmarkStart,Matcher>();
		Map<BookmarkStart,BookmarkEnd> bookmarkpairs = new HashMap<BookmarkStart,BookmarkEnd>();
		Map<BookmarkStart,Run> bookmarkToRunMapper = new HashMap<BookmarkStart,Run>();
		for(Run run : runs.keySet()) {
			Node comment = run.getAncestor(NodeType.COMMENT);
			if(comment!=null) {
				continue;
			}
			Matcher matcher = runs.get(run);
			String format = matcher.group(1).trim();
			String expression = matcher.group(2).trim();
			String value = Objects.toString(ExpressionEvaluator.evaluateExpression(expression, data), "");
			//String value = SimpleValueEvaluator.evaluateEL(expression, data);
			if(format.equalsIgnoreCase("L")||format.equalsIgnoreCase("")||format.equalsIgnoreCase("M")||format.equalsIgnoreCase("R")) {
				
				String bmName = String.format("bm_%s_%d", expression , run.hashCode());
				builder.moveTo(run);
				BookmarkStart start = builder.startBookmark(bmName);
				BookmarkEnd end = builder.endBookmark(bmName);
				run.getParentNode().insertAfter(end, run);
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
