package org.word2pdf.internal;

import com.aspose.words.Run;

public class PlaceHolder {
	
	private Run startNode;
	private Run endNode;
	
	public PlaceHolder(Run startNode, Run endNode) {
		this.startNode = startNode;
		this.endNode = endNode;
	}

	public Run getStartNode() {
		return startNode;
	}

	public void setStartNode(Run startNode) {
		this.startNode = startNode;
	}

	public Run getEndNode() {
		return endNode;
	}

	public void setEndNode(Run endNode) {
		this.endNode = endNode;
	}

}
