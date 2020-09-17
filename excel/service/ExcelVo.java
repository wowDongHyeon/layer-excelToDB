package gmx.gis.excel.service;

import java.util.HashMap;

/**
 * @author 민동현
 *
 */
public class ExcelVo {

	private String tableName;
	private String tableKoreaName;
	private String layerType;
	private String _annox;
	private String _annoy;
	private String point;
	private String line;
	private HashMap<String,Object> columnKrNmMap;
	private HashMap<String,Object> columnTypeMap;
	private HashMap<String,Object> columnValueMap;
	public String getTableName() {
		return tableName;
	}
	public void setTableName(String tableName) {
		this.tableName = tableName;
	}
	public String getTableKoreaName() {
		return tableKoreaName;
	}
	public void setTableKoreaName(String tableKoreaName) {
		this.tableKoreaName = tableKoreaName;
	}
	public String getLayerType() {
		return layerType;
	}
	public void setLayerType(String layerType) {
		this.layerType = layerType;
	}
	public String get_annox() {
		return _annox;
	}
	public void set_annox(String _annox) {
		this._annox = _annox;
	}
	public String get_annoy() {
		return _annoy;
	}
	public void set_annoy(String _annoy) {
		this._annoy = _annoy;
	}
	public String getPoint() {
		return point;
	}
	public void setPoint(String point) {
		this.point = point;
	}
	public String getLine() {
		return line;
	}
	public void setLine(String line) {
		this.line = line;
	}
	public HashMap<String, Object> getColumnKrNmMap() {
		return columnKrNmMap;
	}
	public void setColumnKrNmMap(HashMap<String, Object> columnKrNmMap) {
		this.columnKrNmMap = columnKrNmMap;
	}
	public HashMap<String, Object> getColumnTypeMap() {
		return columnTypeMap;
	}
	public void setColumnTypeMap(HashMap<String, Object> columnTypeMap) {
		this.columnTypeMap = columnTypeMap;
	}
	public HashMap<String, Object> getColumnValueMap() {
		return columnValueMap;
	}
	public void setColumnValueMap(HashMap<String, Object> columnValueMap) {
		this.columnValueMap = columnValueMap;
	}


}
