package org.word2pdf.core;

import java.util.ArrayList;
import java.util.List;

import com.aspose.words.CommentRangeEnd;
import com.aspose.words.CommentRangeStart;
import com.aspose.words.Font;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import com.aspose.words.ReplacingArgs;
import com.aspose.words.Run;
import com.aspose.words.Underline;

public class PlaceHolderResolver {

	public static Run splitRun(Run run, int position) throws Exception {
		Run afterRun = (Run) run.deepClone(true);
		afterRun.setText(run.getText().substring(position));
		run.setText(run.getText().substring((0), (0) + (position)));
		run.getParentNode().insertAfter(afterRun, run);
		return afterRun;
	}

	public static Run resolveRun(ReplacingArgs e) throws Exception {
		ArrayList runs = new ArrayList();
		Node currentNode = e.getMatchNode();

		if (e.getMatchOffset() > 0) {
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
		for (int i = 1; i < runs.size(); i++) {
			Run run = (Run) runs.get(i);
			String appendStr = run.getText();
			ret.setText(ret.getText() + appendStr);
			run.remove();
		}
		return ret;
	}

	public static Run resolveRunWithUnderline(ReplacingArgs e) throws Exception {
		Run currentNode = resolveRun(e);
		return resolveRunWithUnderline(currentNode);
	}

	public static Run resolveRunWithUnderline(Run currentNode) {
		// todo
		Font font = currentNode.getFont();
		if (font.getUnderline() == Underline.NONE) {
			return currentNode;
		}

		List<Run> prevCandidates = new ArrayList<Run>();
		// loop backward
		Run cursorRun = currentNode;
		do {
			Node node = cursorRun.getPreviousSibling();
			if (node == null ) {
				break;
			}
			if(CommentRangeEnd.class.isInstance(node)||CommentRangeStart.class.isInstance(node)) {
				node.remove();
				continue;
			}
			if(!Run.class.isInstance(node)) {
				break;
			}
			Run run = (Run) node;
			if (run.getFont().getUnderline() != font.getUnderline() || run.getText().trim().length() > 0) {
				break;
			}
			prevCandidates.add(run);
			cursorRun = run;
		} while (true);

		List<Run> nextCandidates = new ArrayList<Run>();
		// loop forward
		cursorRun = currentNode;
		do {
			Node node = cursorRun.getNextSibling();
			if (node == null ) {
				break;
			}
			if(CommentRangeEnd.class.isInstance(node)||CommentRangeStart.class.isInstance(node)) {
				node.remove();
				continue;
			}
			if(!Run.class.isInstance(node)) {
				break;
			}
			Run run = (Run) node;
			if (run.getFont().getUnderline() != font.getUnderline() || run.getText().trim().length() > 0) {
				break;
			}
			nextCandidates.add(run);
			cursorRun = run;
		} while (true);

		for (Run run : prevCandidates) {
			currentNode.setText(run.getText() + currentNode.getText());
			run.remove();
		}
		for (Run run : nextCandidates) {
			currentNode.setText(currentNode.getText() + run.getText());
			run.remove();
		}
		return currentNode;
	}
}
