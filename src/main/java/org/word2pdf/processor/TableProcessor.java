package org.word2pdf.processor;

import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.aspose.words.*;
import com.lianmed.evaluator.ExpressionEvaluator;
import com.lianmed.internal.HTable;
import com.lianmed.pdf.WordHandleContext;
import com.lianmed.processor.Processor;
import com.lianmed.util.StringUtil;

public class TableProcessor implements IReplacingCallback, Processor {
	
	private static final Pattern TABLE_ITERATION_DIRECTIVE_PATTERN = Pattern.compile("#(.?)\\{(.*?)\\}");
	
	private WordHandleContext context = null;

	@Override
	public void process(WordHandleContext context) throws Exception {
		this.context = context;
		Map<Table,HTable> htables = retrieveTablesNeedProcess();
		processTables(htables);
	}
	
	@Override
	public int replacing(ReplacingArgs arg0) throws Exception {
		return 0;
	}

	private Map<Table,HTable> retrieveTablesNeedProcess() throws Exception {
		final Map<Table,HTable> htables = new HashMap<Table,HTable>();
		//NodeCollection tables = doc.getChildNodes(NodeType.TABLE, true);
		
		FindReplaceOptions options = new FindReplaceOptions();
		options.setDirection(FindReplaceDirection.FORWARD);
		options.setReplacingCallback(new IReplacingCallback() {

			@Override
			public int replacing(ReplacingArgs target) throws Exception {
				Table tbl = (Table) target.getMatchNode().getAncestor(Table.class);
				if(tbl==null) {
					return 0;
				}
				String format = target.getMatch().group(1).trim();
				Row row = (Row) target.getMatchNode().getAncestor(Row.class);
				Cell cell = (Cell) target.getMatchNode().getAncestor(Cell.class);
				int idx = tbl.getRows().indexOf(row);
				HTable htbl = htables.get(tbl);
				if(htbl==null) {
					htbl = new HTable();
					htables.put(tbl, htbl);
				}
				if(format.equalsIgnoreCase("d")) {
					htbl.setDynamicTable(true);
				}
				if(format.equalsIgnoreCase("i")) {
					htbl.getImgCols().put(row.getCells().indexOf(cell),1);
				}
				if(htbl.isDataError()) {
					return 0;
				}
				if(htbl.getTable()==null) {
					htbl.setTable(tbl);
				}else if(tbl!=htbl.getTable()){
					htbl.setDataError(true);
				}
				if( idx>=0 && (htbl.getBegin()==idx||htbl.getBegin()==-1) ) {
					htbl.setBegin(idx);
				}else {
					htbl.setDataError(true);
				}
				Matcher matcher = target.getMatch();
				htbl.addMata(row.getCells().indexOf(cell), matcher.group(2));
				if(target.getMatchNode() instanceof Run) {
					Run run = (Run) target.getMatchNode();
					htbl.addRun(row.getCells().indexOf(cell), run);
				}
				target.setReplacement("");
				Comment comment = (Comment) target.getMatchNode().getAncestor(Comment.class);
				if(comment!=null) {
					comment.remove();
				}
				return 0;
			}
			
		});
		
		context.getDoc().getRange().replace(TABLE_ITERATION_DIRECTIVE_PATTERN, "", options);
		
		/*for (int i = 0; i < tables.getCount(); i++) {
			Table table = (Table) tables.get(i);
			table.getRange().replace(Pattern.compile("#\\{(.*?)\\}"), "", options);
		}*/
		return htables;
	}
	
	private void processTables(Map<Table,HTable> htables) {
		Object data = context.getData();
		Document doc = context.getDoc();
		DocumentBuilder builder = context.getDocBuilder();
		if(htables==null) {
			return ;
		}
		
		for(HTable htbl : htables.values()) {
			if(htbl.isDataError()) {
				continue;
			}
			Table table = htbl.getTable();
			int numOfCol = table.getRows().get(htbl.getBegin()).getCount();
			
			int size = ExpressionEvaluator.getSizeOfIterationData((String)htbl.getMata().values().toArray()[0], data);
			//int size = SimpleValueEvaluator.getSizeOfIterationData((String)htbl.getMata().values().toArray()[0], data);
			if(size !=-1&&htbl.isDynamicTable()) {
				int num = size-table.getRows().getCount()+htbl.getBegin();
				for(int i=0; i<num; i++) {
					Node nrow = table.getRows().get(htbl.getBegin()+i).deepClone(true);
					table.appendChild(nrow);
				}
			}
			int toi = 0;
			for(int i=htbl.getBegin(); i<table.getRows().getCount(); i++) {
				Row row = table.getRows().get(i);
				Node parent = row.getParentNode();
				if(row.getCells().getCount()!=numOfCol) {
					break;
				}
				int increaseNum = -1;
				for(int j : htbl.getMata().keySet()) {
					Cell cell = row.getCells().get(j);
					
					String expression = htbl.getMata().get(j);
					
					if(StringUtil.search(expression, "idx")>1) {
						if(increaseNum==-1) {
							increaseNum = recurseMergeTable(expression, htbl.getMata(),table,row,i-htbl.getBegin()-toi,j);
						}
						
						for(int k=0;k<increaseNum;k++) {
							if(htbl.getImgCols().containsKey(j)) {
								String sexpression = expression.replaceFirst("idx", i-htbl.getBegin()-toi+"").replaceFirst("idx", k+"");
								Object value = ExpressionEvaluator.evaluateExpression(sexpression, data);
								cell = table.getRows().get(i+k).getCells().get(j);
								Paragraph paragraph = (Paragraph) cell.getChildNodes(NodeType.PARAGRAPH, true).get(0);
								builder.moveTo(paragraph);
								if(value instanceof String) {
									try {
										builder.insertImage((String)value);
									} catch (Exception e) {
									}
								} else if (value instanceof byte[]) {
									try {
										builder.insertImage((byte[]) value);
									} catch (Exception e) {

									}
								} else if(value instanceof InputStream) {
									try {
										builder.insertImage((InputStream) value);
									} catch (Exception e) {

									}
								}
							} else {
								String sexpression = expression.replaceFirst("idx", i-htbl.getBegin()-toi+"").replaceFirst("idx", k+"");
								String value = Objects.toString(ExpressionEvaluator.evaluateExpression(sexpression, data), "");
								Run newChild = null;
								if(htbl.getRuns().containsKey(j)) {
									newChild = (Run) htbl.getRuns().get(j).deepClone(true);
									newChild.setText(value);
								}else {
									newChild = new Run(doc, value);
								}
								cell = table.getRows().get(i+k).getCells().get(j);
								Paragraph paragraph = (Paragraph) cell.getChildNodes(NodeType.PARAGRAPH, true).get(0);
								paragraph.appendChild(newChild);
							}

						}
						
					}else {
						if(htbl.getImgCols().containsKey(j)) {
							Object value = ExpressionEvaluator.evaluateExpression(expression, data, i - htbl.getBegin() - toi);
							Paragraph paragraph = (Paragraph) cell.getChildNodes(NodeType.PARAGRAPH, true).get(0);
							builder.moveTo(paragraph);
							if(value instanceof String) {
								try {
									builder.insertImage((String)value);
								} catch (Exception e) {
								}
							} else if (value instanceof byte[]) {
								try {
									builder.insertImage((byte[]) value);
								} catch (Exception e) {

								}
							} else if(value instanceof InputStream) {
								try {
									builder.insertImage((InputStream) value);
								} catch (Exception e) {

								}
							}


						} else {
							String value = Objects.toString(ExpressionEvaluator.evaluateExpression(expression, data, i-htbl.getBegin()-toi), "");

							Run newChild = null;
							if(htbl.getRuns().containsKey(j)) {
								newChild = (Run) htbl.getRuns().get(j).deepClone(true);
								newChild.setText(value);
							}else {
								newChild = new Run(doc, value);
							}

							Paragraph paragraph = (Paragraph) cell.getChildNodes(NodeType.PARAGRAPH, true).get(0);
							paragraph.appendChild(newChild);
						}
					}
					
					
				}
				if(increaseNum > 0) {
					i = i + increaseNum -1;
					toi = toi + increaseNum -1;
					increaseNum = -1;
				}
				
			}
			
			if(context.getListener()!=null) {
				Shape shape = (Shape) table.getAncestor(NodeType.SHAPE);
				if(shape!=null) {
					
					Node[] nodes = shape.getChildNodes().toArray();
					for(Node node : nodes) {
						if(node.getNodeType()==NodeType.PARAGRAPH) {
							Paragraph para = (Paragraph) node;
							if(para.getText().trim().length()>0) {
								continue;
							}
							para.getParagraphFormat().setLineSpacingRule(LineSpacingRule.EXACTLY);
							para.getParagraphFormat().setLineSpacing(1);
						}
					}
					try {
						doc.updatePageLayout();
						LayoutCollector collector = context.getCollector();
						LayoutEnumerator enumerator = context.getEnumerator();
						enumerator.setCurrent(collector.getEntity(shape));
						double origin = enumerator.getRectangle().getHeight();
						shape.getTextBox().setFitShapeToText(true);
						doc.updatePageLayout();
						enumerator.setCurrent(collector.getEntity(shape));
						double adaptive = enumerator.getRectangle().getHeight();
						context.getListener().fitShapeToText(origin, adaptive);
						shape.getTextBox().setFitShapeToText(false);
						doc.updatePageLayout();
					} catch (Exception e) {
						
					}
				}
			}
			
		}
	}

	private int recurseMergeTable(String expression,Map<Integer,String> mataData,Table table, Row row, int rowIdx, int colIdx) {
		
		/*int offset = StringUtil.indexOf(expression, "idx", 1);
		String subExp = expression.substring(0, offset-1);*/
		expression = expression.replaceFirst("idx", rowIdx+"");
		int size = ExpressionEvaluator.getSizeOfIterationData(expression, context.getData());
		if(size<1) {
			return size;
		}
		Row newRow = (Row)row.deepClone(true);
		for(int i=0; i<row.getCells().getCount(); i++) {
			Cell cell = (Cell) row.getCells().get(i);
			if(StringUtil.search(mataData.get(i), "idx")==1) {
				cell.getCellFormat().setVerticalMerge(CellMerge.FIRST);
			}
			if(i<colIdx) {
				
			}
			
		}
		for(int i=0; i<newRow.getCells().getCount(); i++) {
			Cell cell = (Cell) newRow.getCells().get(i);
			if(StringUtil.search(mataData.get(i), "idx")==1) {
				cell.getCellFormat().setVerticalMerge(CellMerge.PREVIOUS);
			}
			cell.getFirstParagraph().removeAllChildren();
			
		}
		for(int i=1;i<size;i++) {
			table.insertAfter(newRow.deepClone(true), row);
		}
		return size;
	}

}
