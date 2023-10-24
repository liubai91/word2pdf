package org.word2pdf.util;

import java.util.ArrayList;
import java.util.List;

import com.aspose.words.BookmarkEnd;
import com.aspose.words.BookmarkStart;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Font;
import com.aspose.words.LayoutCollector;
import com.aspose.words.LayoutEnumerator;
import com.aspose.words.LineSpacingRule;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.ReplacingArgs;
import com.aspose.words.Run;
import com.aspose.words.Section;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import com.aspose.words.TextBox;
import com.aspose.words.Underline;
import com.aspose.words.WrapType;

public class NodeUtil {
	
	public static Shape insertTextBoxCoverBookmark(DocumentBuilder builder, LayoutCollector collector, LayoutEnumerator enumerator,
			BookmarkStart bmStart, BookmarkEnd bmEnd) throws Exception {
		Section section = (Section) bmStart.getAncestor(NodeType.SECTION);
		builder.moveTo(bmStart);
		
		
		
		enumerator.setCurrent(collector.getEntity(bmStart));
		
		
		
		double left = enumerator.getRectangle().getX();
		double top = enumerator.getRectangle().getY();
		double height = enumerator.getRectangle().getHeight();
		Node table = bmStart.getAncestor(NodeType.TABLE);
		Shape shape = (Shape) bmStart.getAncestor(NodeType.SHAPE);
		if(table!=null||shape!=null) {
			Document doc = collector.getDocument();
			int pageidx = enumerator.getPageIndex();
			Node[] nodes = doc.getChildNodes(NodeType.PARAGRAPH, true).toArray();
			for(Node node : nodes) {
				if(node.getAncestor(NodeType.TABLE)!=null) {
					continue;
				}
				if(node.getAncestor(NodeType.SHAPE)!=null) {
					continue;
				}
				if(collector.getEntity(node)==null) {
					continue;
				}
				enumerator.setCurrent(collector.getEntity(node));
				if(enumerator.getPageIndex()==pageidx) {
					builder.moveTo(node);
					break;
				}
			}
		}
		enumerator.setCurrent(collector.getEntity(bmEnd));
		double width = enumerator.getRectangle().getX() - left;
		if(shape!=null&&shape.getShapeType()==ShapeType.TEXT_BOX) {
			left += shape.getTextBox().getInternalMarginLeft();
			top += shape.getTextBox().getInternalMarginTop();
			width += shape.getTextBox().getInternalMarginLeft();
		}
		double maxWidth = section.getPageSetup().getPageWidth() - section.getPageSetup().getRightMargin() - left;
		width = width>maxWidth ? maxWidth: width;
		
		Shape textBoxShape = builder.insertShape(ShapeType.TEXT_BOX, RelativeHorizontalPosition.PAGE, left, 
				RelativeHorizontalPosition.PAGE, top, width, height, WrapType.NONE);
		textBoxShape.setZOrder(1000);
		textBoxShape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
		textBoxShape.setStroked(false);
		TextBox textBox = textBoxShape.getTextBox();
		textBox.setInternalMarginBottom(0);
		textBox.setInternalMarginLeft(0);
		textBox.setInternalMarginRight(0);
		textBox.setInternalMarginTop(0);
		textBox.setFitShapeToText(false);
		Paragraph paragraph = textBoxShape.getLastParagraph();
		paragraph.getParagraphFormat().setLineSpacingRule(LineSpacingRule.EXACTLY);
		paragraph.getParagraphFormat().setLineSpacing(height);
		return textBoxShape;
	}
	
	public static Run splitRun(Run run, int position) throws Exception {
		Run afterRun = (Run) run.deepClone(true);
		afterRun.setText(run.getText().substring(position));
		run.setText(run.getText().substring((0), (0) + (position)));
		run.getParentNode().insertAfter(afterRun, run);
		return afterRun;
	}
	
	public static Run extractMatchedRun(ReplacingArgs e) throws Exception {
		ArrayList runs = new ArrayList();
		Node currentNode =  e.getMatchNode();
		
		if(e.getMatchOffset()>0) {
			currentNode = splitRun((Run) e.getMatchNode(), e.getMatchOffset());
		}
		int remainingLength = e.getMatch().group().length();
		while ((remainingLength > 0) && (currentNode != null) && (currentNode.getText().length() <= remainingLength)) {
			runs.add(currentNode);
			remainingLength = remainingLength - currentNode.getText().length();

			// Select the next Run node.
			// Have to loop because there could be other nodes such as BookmarkStart etc.
			do {
				currentNode = currentNode.getNextSibling();
			} while ((currentNode != null) && (currentNode.getNodeType() != NodeType.RUN));
		}
		
		if ((currentNode != null) && (remainingLength > 0)) {
			splitRun((Run) currentNode, remainingLength);
			runs.add(currentNode);
		}
		Run ret = (Run) runs.get(0);
		for(int i=1; i<runs.size();i++) {
			Run run = (Run)runs.get(i);
			String appendStr = run.getText();
			ret.setText(ret.getText()+appendStr);
			run.remove();
		}
		return ret;
	}
	
	public static Run extractMatchedRunWithUnderline(ReplacingArgs e) throws Exception {
		
		Run currentNode = extractMatchedRun(e);
		//todo
		Font font = currentNode.getFont();
		if(font.getUnderline()==Underline.NONE) {
			return currentNode;
		}
		
		List<Run> prevCandidates = new ArrayList<Run>();
		//loop backward
		Run cursorRun = currentNode;
		do {
			Node node = cursorRun.getPreviousSibling();
			if(node==null||!Run.class.isInstance(node)) {
				break;
			}
			Run run = (Run) node;
			if(run.getFont().getUnderline()!=font.getUnderline()||run.getText().trim().length()>0) {
				break;
			}
			prevCandidates.add(run);
			cursorRun = run;
		}while(true);
		
		List<Run> nextCandidates = new ArrayList<Run>();
		//loop forward
		cursorRun = currentNode;
		do {
			Node node = cursorRun.getNextSibling();
			if(node==null||!Run.class.isInstance(node)) {
				break;
			}
			Run run = (Run) node;
			if(run.getFont().getUnderline()!=font.getUnderline()||run.getText().trim().length()>0) {
				break;
			}
			nextCandidates.add(run);
			cursorRun = run;
		}while(true);
		
		for(Run run : prevCandidates) {
			currentNode.setText(run.getText()+currentNode.getText());
			run.remove();
		}
		for(Run run: nextCandidates) {
			currentNode.setText(currentNode.getText()+run.getText());
			run.remove();
		}
		return currentNode;
	}

}
