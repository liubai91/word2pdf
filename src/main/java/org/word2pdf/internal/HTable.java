package org.word2pdf.internal;

import java.util.HashMap;
import java.util.Map;

import com.aspose.words.Run;
import com.aspose.words.Table;

public class HTable {
	
	private Table table;
	private int begin = -1;
	private Map<Integer,String> mata = new HashMap<Integer,String>();
	private Map<Integer,Run> runs = new HashMap<Integer,Run>();

	private Map<Integer,Object> imgCols = new HashMap<>();
	private boolean dataError = false;
	private boolean dynamicTable = false;
	
	
	public Table getTable() {
		return table;
	}
	public void setTable(Table table) {
		this.table = table;
	}
	public int getBegin() {
		return begin;
	}
	public void setBegin(int begin) {
		this.begin = begin;
	}
	public Map<Integer, String> getMata() {
		return mata;
	}
	public void setMata(Map<Integer, String> mata) {
		this.mata = mata;
	}
	public boolean isDataError() {
		return dataError;
	}
	public void setDataError(boolean dataError) {
		this.dataError = dataError;
	}
	public boolean isDynamicTable() {
		return dynamicTable;
	}
	public void setDynamicTable(boolean dynamicTable) {
		this.dynamicTable = dynamicTable;
	}
	public void addMata(Integer idx,String expression) {
		mata.put(idx, expression);
	}
	public Map<Integer, Run> getRuns() {
		return runs;
	}
	public void setRuns(Map<Integer, Run> runs) {
		this.runs = runs;
	}
	
	public void addRun(Integer idx,Run run) {
		runs.put(idx, run);
	}

	public Map<Integer, Object> getImgCols() {
		return imgCols;
	}

	public void setImgCols(Map<Integer, Object> imgCols) {
		this.imgCols = imgCols;
	}
}
