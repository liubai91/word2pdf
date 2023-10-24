package org.word2pdf.processor;

import com.aspose.words.Document;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.lianmed.pdf.WordHandleContext;
import com.lianmed.pdf.WordHandler;
import com.lianmed.processor.Processor;

public class IndentProcessor implements Processor {

	@Override
	public void process(WordHandleContext context) throws Exception {
		Document doc = context.getDoc();
		Node[] paras = doc.getChildNodes(NodeType.PARAGRAPH, true).toArray();
		for(Node node : paras) {
			Paragraph para = (Paragraph) node;
			para.getParagraphFormat().setLeftIndent(0f);
		}
	}

}
