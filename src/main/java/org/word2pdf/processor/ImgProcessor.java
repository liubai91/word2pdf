package org.word2pdf.processor;

import java.io.InputStream;
import java.util.regex.Pattern;

import com.aspose.words.DocumentBuilder;
import com.aspose.words.FindReplaceDirection;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import com.lianmed.evaluator.ExpressionEvaluator;
import com.lianmed.pdf.WordHandleContext;
import com.lianmed.processor.Processor;

public class ImgProcessor implements Processor, IReplacingCallback {

	private static final Pattern IMG_DIRECTIVE_PATTERN = Pattern.compile("(<<image.*?\\[)(.*)(\\].*?>>)");

	
	private WordHandleContext context;

	@Override
	public void process(WordHandleContext context) throws Exception {
		FindReplaceOptions options = new FindReplaceOptions();
		options.setDirection(FindReplaceDirection.FORWARD);
		options.setReplacingCallback(this);
		context.getDoc().getRange().replace(IMG_DIRECTIVE_PATTERN, "", options);
	}

	@Override
	public int replacing(ReplacingArgs e) throws Exception {
		Object data = context.getData();
		DocumentBuilder builder = context.getDocBuilder();
		Object value = ExpressionEvaluator.evaluateExpression(e.getMatch().group(2), data);
		//Object value = SimpleEvaluator.getProperty(e.getMatch().group(2), data, -1);
		if (value == null || !InputStream.class.isInstance(value)) {
			e.setReplacement("");
			return ReplaceAction.REPLACE;
		}
		InputStream stream = (InputStream) value;
		Node matchedNode = e.getMatchNode();
		builder.moveTo(matchedNode);
		Shape img = builder.insertImage(stream);
		Shape textboxShape = (Shape) img.getAncestor(NodeType.SHAPE);
		if (textboxShape != null && textboxShape.getShapeType() == ShapeType.TEXT_BOX) {
			textboxShape.getTextBox().setInternalMarginLeft(0);
			textboxShape.getTextBox().setInternalMarginTop(0);

			img.setWidth(textboxShape.getWidth());
			img.setHeight(textboxShape.getHeight());

		}
		e.setReplacement("");
		return 0;
	}

}
