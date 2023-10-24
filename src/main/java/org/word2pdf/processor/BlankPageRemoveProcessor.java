package org.word2pdf.processor;

import java.util.ArrayList;

import com.aspose.words.Document;
import com.aspose.words.LayoutCollector;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.SaveFormat;
import com.aspose.words.SectionCollection;
import com.lianmed.pdf.WordHandleContext;
import com.lianmed.pdf.WordHandler;
import com.lianmed.processor.Processor;

public class BlankPageRemoveProcessor implements Processor {
	
	private WordHandleContext context;
	

	@Override
	public void process(WordHandleContext context) throws Exception {
		this.context = context;
		deleteBlankPage();
	}
	
	private void deleteBlankPage() throws Exception {
		Document doc = context.getDoc();
		SectionCollection sections = doc.getSections();
		for(int i=0;i<sections.getCount();i++) {
			if(sections.get(i).toString(SaveFormat.TEXT).trim().equals("")) {
				sections.get(i).remove();
			}
		}
		String pageText = "";
		LayoutCollector lc = new LayoutCollector(doc);
		if(doc.getLastSection()==null||doc.getLastSection().getBody()==null) {
			return;
		}
		int pages = lc.getStartPageIndex(doc.getLastSection().getBody().getLastParagraph());
		for (int i = 1; i <= pages; i++) {
			ArrayList<Paragraph> nodes = getNodesByPage(i);
			for(Paragraph p : nodes) {
				pageText += p.toString(SaveFormat.TEXT).trim();
			}
			if (pageText.trim().equals("")){
				while(nodes.size()>0) {
					nodes.remove(0);
				}
			}
			pageText = "";
		}
	}
	
	private ArrayList<Paragraph> getNodesByPage(int page) throws Exception{
		Document doc = context.getDoc();
		ArrayList<Paragraph> nodes = new ArrayList<Paragraph>();
		LayoutCollector lc = new LayoutCollector(doc);
		NodeCollection collections = doc.getChildNodes(NodeType.PARAGRAPH, true);
		for(int i=0;i<collections.getCount();i++) {
			Paragraph para = (Paragraph) collections.get(i);
			if (lc.getStartPageIndex(para) == page || para.isEndOfSection())
				nodes.add(para);
		}
		
		return nodes;
	}

}
