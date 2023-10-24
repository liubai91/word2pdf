package org.word2pdf.processor;

import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.aspose.words.FindReplaceDirection;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.lianmed.evaluator.ExpressionEvaluator;
import com.lianmed.pdf.WordHandleContext;
import com.lianmed.processor.Processor;

public class ConditionalBlockProcessor implements Processor,IReplacingCallback {
	
	private static final Pattern CHCKEBOX_DIRECTIVE_PATTERN = Pattern.compile("(<<\\s*if\\s*\\[)(.*?)(\\]\\s*>>)");
	
	private WordHandleContext context;

	@Override
	public void process(WordHandleContext context) throws Exception {
		this.context = context;
		FindReplaceOptions options = new FindReplaceOptions();
		options.setDirection(FindReplaceDirection.BACKWARD);
		options.setReplacingCallback(this);
		context.getDoc().getRange().replace(CHCKEBOX_DIRECTIVE_PATTERN, "", options);

	}

	@Override
	public int replacing(ReplacingArgs e) throws Exception {
		Matcher matcher = e.getMatch();
		String expression = matcher.group(2);
		//boolean checked = ExpressionEvaluator.evaluateEL(expression, context.getData());
		boolean checked = false;
		try {
			checked = (boolean) ExpressionEvaluator.evaluateExpression(expression, context.getData());
		} catch (Exception ex) {
			
		}
		try {
			int i = Integer.valueOf(Objects.toString(ExpressionEvaluator.evaluateExpression(expression, context.getData()),""));
			checked = i > 0 ? true : false;
		} catch (Exception ex) {
			
		}
		String txt = matcher.group();
		if(checked) {
			e.setReplacement("<<if [true]>>");
		}else {
			e.setReplacement("<<if [false]>>");
		}
		
		return ReplaceAction.REPLACE;
	}

}
