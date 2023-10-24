package org.word2pdf.processor;

import java.io.UnsupportedEncodingException;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.lianmed.processor.Processor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.aspose.words.BookmarkEnd;
import com.aspose.words.BookmarkStart;
import com.aspose.words.Cell;
import com.aspose.words.Comment;
import com.aspose.words.CommentRangeEnd;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.LayoutCollector;
import com.aspose.words.LayoutEnumerator;
import com.aspose.words.Node;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.Run;
import com.aspose.words.Shape;
import com.lianmed.core.PlaceHolderResolver;
import com.lianmed.evaluator.ExpressionEvaluator;
import com.lianmed.pdf.WordHandleContext;
import com.lianmed.text.TextAlign;
import com.lianmed.util.HexUtil;
import com.lianmed.util.NodeUtil;

/**
 * 批注处理器
 * @author nhl220
 *
 * 仅处理下面两种情况：
 * 1.批注前一个节点是个未选中的复选框
 * 2.批注位于带下划线的占位符环境中
 * 其他情况一概不再做处理，仅删除批注
 */
public class CommentProcessor implements Processor {
	
	private final Logger log = LoggerFactory.getLogger(CommentProcessor.class);
	
	private static final Pattern CHCKEBOX_DIRECTIVE_PATTERN = Pattern.compile("@\\{(.*?)\\}");
	private static final Pattern DATA_ACCESS_DIRECTIVE_PATTERN = Pattern.compile("\\$(.?)\\{(.*?)\\}");
	
	final private static char[] CHECK = {0xFE};
	final private static char[] UNCHECK = {0xA8};
	final private static String CHECK_HEX = "EF83BE";
	final private static String UNCHECK_HEX = "EF82A8";
	
	private WordHandleContext context;
	
	private Map<Run,Matcher> run2MatcherMapper;
	private Map<Run,Run> run2RunMapper;
	

	@Override
	public void process(WordHandleContext context) throws Exception {
		this.context = context;
		run2MatcherMapper = new HashMap<Run,Matcher>();
		run2RunMapper = new HashMap<Run,Run>();
		
		Node[] nodes = context.getDoc().getChildNodes(NodeType.COMMENT, true).toArray();
		for(Node node : nodes) {
			Comment comment = (Comment) node;
			String content = comment.getText().trim();
			Matcher matcher = DATA_ACCESS_DIRECTIVE_PATTERN.matcher(content);
			if(matcher.matches()) {
				preHandlePlaceHolder(comment, matcher);
			}else if((matcher = CHCKEBOX_DIRECTIVE_PATTERN.matcher(content)).matches())  {
				handleCheckBox(comment, matcher);
			}
			if(comment.getParentNode()!=null) {
				comment.remove();
			}
		}
		doHandlePlaceHolder();
		/*for(Node node : nodes) {
			node.remove();
		}*/
	}

	private void doHandlePlaceHolder() throws Exception {
		Object data = context.getData();
		Document doc = context.getDoc();
		DocumentBuilder builder = context.getDocBuilder();
		/*LayoutCollector collector = new LayoutCollector(doc);
		LayoutEnumerator enumerator = new LayoutEnumerator(doc);*/
		LayoutCollector collector = context.getCollector();
		LayoutEnumerator enumerator = context.getEnumerator();
		Map<BookmarkStart,Matcher> bookmarks = new HashMap<BookmarkStart,Matcher>();
		Map<BookmarkStart,BookmarkEnd> bookmarkpairs = new HashMap<BookmarkStart,BookmarkEnd>();
		Map<BookmarkStart,Run> bookmarkToRunMapper = new HashMap<BookmarkStart,Run>();
		for(Run run : run2MatcherMapper.keySet()) {
			Matcher matcher = run2MatcherMapper.get(run);
			String format = matcher.group(1).trim();
			String expression = matcher.group(2).trim();
			String value = Objects.toString(ExpressionEvaluator.evaluateExpression(expression, data), "");
			//String value = SimpleValueEvaluator.evaluateEL(expression, data);
			if(format.equalsIgnoreCase("L")||format.equalsIgnoreCase("")||format.equalsIgnoreCase("M")||format.equalsIgnoreCase("R")) {
				builder.moveTo(run);
				String bmName = String.format("bm_%s_%d", expression , run.hashCode());
				BookmarkStart start = builder.startBookmark(bmName);
				BookmarkEnd end = builder.endBookmark(bmName);
				run.getParentNode().insertAfter(end, run);
				bookmarks.put(start, matcher);
				bookmarkpairs.put(start, end);
				bookmarkToRunMapper.put(start, run);
			}else {
				Run destStyle = (Run) run2RunMapper.get(run).deepClone(true);
				destStyle.setText(value);
				run.getParentNode().insertAfter(destStyle, run);
				run.remove();
				destStyle.getFont().setUnderline(run.getFont().getUnderline());
				//run.setText(value);
			}
		}
		
		doc.updatePageLayout();
		
		for(BookmarkStart bmStart : bookmarks.keySet()) {
			Shape textBoxShape = NodeUtil.insertTextBoxCoverBookmark(builder, collector, enumerator, bmStart, bookmarkpairs.get(bmStart));
			Paragraph paragraph = textBoxShape.getLastParagraph();
			paragraph.getParagraphFormat().setAlignment(ParagraphAlignment.LEFT);
			Run ori = run2RunMapper.get(bookmarkToRunMapper.get(bmStart));
			Run run = (Run)  ori.deepClone(false);
			run.getFont().setName(ori.getFont().getName());
			run.getFont().setSize(ori.getFont().getSize());
			run.getFont().setBold(ori.getFont().getBold());
			run.getFont().setItalic(ori.getFont().getItalic());
			paragraph.getParagraphFormat().getStyle().getFont().getName();
			//Run run = (Run) bookmarkToRunMapper.get(bmStart).deepClone(true);
			Matcher matcher = bookmarks.get(bmStart);
			String format = matcher.group(1).trim();
			String expression = matcher.group(2).trim();
			String value = Objects.toString(ExpressionEvaluator.evaluateExpression(expression, data), "");
			//String value = SimpleValueEvaluator.evaluateEL(expression, data);
			//value = expression;
			/*if(value.trim().length()==0) {
				value = expression;
			}*/
			String origin = run.getText();
			
			run.setText(TextAlign.formatText(TextAlign.LEFT, origin, value, run.getFont()));
			//run.setText(value);
			run.getFont().setUnderline(bookmarkToRunMapper.get(bmStart).getFont().getUnderline());
			paragraph.appendChild(run);
		}
	}

	private void preHandlePlaceHolder(Comment comment, Matcher matcher) {
		Node prevNode = comment.getPreviousSibling();
		
		if(handleTableContext(comment, matcher)) {
			return;
		}
		if( prevNode==null || (!Run.class.isInstance(prevNode) && !CommentRangeEnd.class.isInstance(prevNode))) {
			return ;
		}
		Run prevRun = null;
		if(!Run.class.isInstance(prevNode)) {
			prevRun = (Run) prevNode.getPreviousSibling();
		}else {
			prevRun = (Run) prevNode;
		}
		
		/*if(prevRun.getFont().getUnderline()==Underline.NONE) {
			return ;
		}*/
		/*Node nextNode = comment.getNextSibling();
		if(nextNode!=null || !Run.class.isInstance(nextNode)) {
			Run nextRun = (Run) nextNode;
			if(nextRun.getFont().getUnderline()==prevRun.getFont().getUnderline()) {
			//if(nextRun.getFont().getUnderline()!=Underline.NONE) {
				prevRun.setText(prevRun.getText()+nextRun.getText());
				nextRun.remove();
			}
		}*/
		NodeCollection runs = comment.getChildNodes(NodeType.RUN, true);
		if(runs.getCount()==0) {
			return ;
		}
		comment.remove();
		prevRun = PlaceHolderResolver.resolveRunWithUnderline(prevRun);
		Run run = (Run) runs.get(0);
		run2MatcherMapper.put(prevRun, matcher);
		run2RunMapper.put(prevRun, run);
	}

	private boolean handleTableContext(Comment comment, Matcher matcher) {
		Document doc = context.getDoc();
		Object data = context.getData();
		Cell cell = (Cell) comment.getAncestor(NodeType.CELL);
		if(cell == null||!cell.getText().trim().equals(matcher.group())) {
			return false;
		}
		NodeCollection tables = cell.getChildNodes(NodeType.TABLE, true);
		if(tables.getCount()>0) {
			return false;
		}
		Node node = cell.getFirstChild();
		
		if(Paragraph.class.isInstance(node)) {
			Paragraph para = (Paragraph) node;
			String expression = matcher.group(2).trim();
			String value = Objects.toString(ExpressionEvaluator.evaluateExpression(expression, data), "");
			//String value = SimpleValueEvaluator.evaluateEL(expression, data);
			Run newChild = null;
			NodeCollection runs = comment.getChildNodes(NodeType.RUN, true);
			if(runs.getCount()==0) {
				newChild = new Run(doc, Objects.toString(value));
			}else {
				newChild = (Run) runs.get(0).deepClone(true);
				newChild.setText(Objects.toString(value));
			}
			para.appendChild(newChild);
		}
		return true;
	}

	private void handleCheckBox(Comment comment, Matcher matcher) throws UnsupportedEncodingException {
		Object data = context.getData();
		Node prevNode = comment.getPreviousSibling();
		if( prevNode==null || !Run.class.isInstance(prevNode) ) {
			return ;
		}
		Run prevRun = (Run) prevNode;
		String hex = HexUtil.encodeHexString(prevRun.getText().getBytes("utf-8"), false);
		if(hex.equals(CHECK_HEX) || hex.equals(UNCHECK_HEX)) {
			boolean checked = false;
			try {
				checked = (boolean) ExpressionEvaluator.evaluateExpression(matcher.group(1), data);
			} catch (Exception e) {
				
			}
			try {
				int i = Integer.valueOf(Objects.toString(ExpressionEvaluator.evaluateExpression(matcher.group(1), data),""));
				checked = i > 0 ? true : false;
			} catch (Exception e) {
				
			}
			String txt = checked ? new String(CHECK) : new String(UNCHECK);
			prevRun.setText(new String(txt));
		}
	}

}
