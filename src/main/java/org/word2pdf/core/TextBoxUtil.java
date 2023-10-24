package org.word2pdf.core;

import com.aspose.words.BookmarkEnd;
import com.aspose.words.BookmarkStart;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.LayoutCollector;
import com.aspose.words.LayoutEnumerator;
import com.aspose.words.LineSpacingRule;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.Section;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import com.aspose.words.TextBox;
import com.aspose.words.WrapType;

public class TextBoxUtil {
	
	public static Shape insertTextBoxCoverBookmark(DocumentBuilder builder, LayoutCollector collector, LayoutEnumerator enumerator,
			BookmarkStart bmStart, BookmarkEnd bmEnd) throws Exception {
		Section section = (Section) bmStart.getAncestor(NodeType.SECTION);
		builder.moveTo(bmStart);
		enumerator.setCurrent(collector.getEntity(bmStart));
		double left = enumerator.getRectangle().getX();
		double top = enumerator.getRectangle().getY();
		double height = enumerator.getRectangle().getHeight();
		enumerator.setCurrent(collector.getEntity(bmEnd));
		double width = enumerator.getRectangle().getX() - left;
		double maxWidth = section.getPageSetup().getPageWidth() - section.getPageSetup().getRightMargin() - left;
		width = width>maxWidth ? maxWidth: width;
		Shape textBoxShape = builder.insertShape(ShapeType.TEXT_BOX, RelativeHorizontalPosition.PAGE, left, 
				RelativeHorizontalPosition.PAGE, top, width, height, WrapType.NONE);
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

}
