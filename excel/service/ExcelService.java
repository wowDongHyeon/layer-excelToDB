package gmx.gis.excel.service;



import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.InvocationTargetException;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Pattern;

import javax.annotation.Resource;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.PlatformTransactionManager;
import org.springframework.validation.Validator;
import org.springframework.web.bind.WebDataBinder;
import org.springframework.web.bind.annotation.InitBinder;
import org.springframework.web.multipart.MultipartFile;

import egovframework.rte.fdl.cmmn.EgovAbstractServiceImpl;
import gmx.gis.layer.service.LayerGroupService;
import gmx.gis.layer.service.LayerGroupVo;
import gmx.gis.layer.service.LayerService;
import gmx.gis.layer.service.LayerStyleService;
import gmx.gis.layer.service.LayerStyleVo;
import gmx.gis.layer.service.LayerVo;
import gmx.gis.excel.service.ExcelVo;
import gmx.gis.sysmgr.service.ColumnKoreaNameVo;
import gmx.gis.sysmgr.service.ColumnNameVo;
import gmx.gis.sysmgr.service.ColumnTypeVo;
import gmx.gis.sysmgr.service.ColumnLengthVo;
import gmx.gis.util.code.ReflectionUtil;


/**
 * @author 민동현
 *
 */
@Service("excelService")
public class ExcelService extends EgovAbstractServiceImpl {



    @Autowired private LayerService svc;

	@Autowired private LayerGroupService group;

	@Autowired private LayerStyleService style;

	@Resource(name = "excelMapper")
    private ExcelMapper mapper;

	@Resource(name = "txManager")
   	PlatformTransactionManager transactionManager;

    @Resource
    private Validator validator;


    @InitBinder
    private void initBinder(WebDataBinder binder){
        binder.setValidator(this.validator);

    }

    private static final String EXCEL_SCHEMA = "excel";

    private String tblNm;
    private String tblKrNm;
    private String lyrTyp;
    private String errorMsg;

    private int colCnt;
    private ColumnNameVo colNmObj;
    private ColumnKoreaNameVo colKrNmObj;
    private ColumnTypeVo colTypeObj;
    private ColumnLengthVo colLenObj;

	private ReflectionUtil<ColumnNameVo> refColNm;
    private ReflectionUtil<ColumnKoreaNameVo> refColKrNm;
    private ReflectionUtil<ColumnTypeVo> refColType;
    private ReflectionUtil<ColumnLengthVo> refColLen;


    private void init() throws Exception {
    	colCnt=2;
    	errorMsg="";
    	colNmObj = new ColumnNameVo();
    	colKrNmObj = new ColumnKoreaNameVo();
    	colTypeObj = new ColumnTypeVo();
	    colLenObj = new ColumnLengthVo();

	    refColNm = new ReflectionUtil<ColumnNameVo>(colNmObj);
	    refColKrNm = new ReflectionUtil<ColumnKoreaNameVo>(colKrNmObj);
	    refColType = new ReflectionUtil<ColumnTypeVo>(colTypeObj);
	    refColLen = new ReflectionUtil<ColumnLengthVo>(colLenObj);

    }

	public List<String> getExcelColumnList(MultipartFile file) throws Exception {

		List<String> list = null;
		String[] splitList=file.getOriginalFilename().split(Pattern.quote("."));
		String excelType=splitList[splitList.length-1];
	    if("xlsx".equals(excelType)){
	    	list=getExcelColumnListAtXlsx(file);
	    }
	    else if("xls".equals(excelType)){
	    	list=getExcelColumnListAtXls(file);
	    }else{
	    	throw new Exception();
        }

		return list;

	}

	private List<String> getExcelColumnListAtXlsx(MultipartFile file) throws IOException {
		Workbook workbook=null;
		List<String> list=new ArrayList<String>();
		try{
			workbook= new XSSFWorkbook(file.getInputStream());
	        Sheet sheet = workbook.getSheetAt(0);
	        Iterator<Row> rows = sheet.iterator();

	        while (rows.hasNext()) {
		       	Row currentRow = rows.next();
		       	Iterator<Cell> cellsInRow = currentRow.iterator();
		       	while (cellsInRow.hasNext()) {
		       		Cell currentCell = cellsInRow.next();
		       		list.add(currentCell.getStringCellValue());
		       	}
		       	break;
	        }
		}catch(Exception e){
			list=null;
			e.printStackTrace();
		}finally{
			 if(workbook!=null){
				 workbook.close();
			 }
		}
		return list;
	}

    private List<String> getExcelColumnListAtXls(MultipartFile file) throws IOException {
    	HSSFWorkbook workbook=null;
		List<String> list=new ArrayList<String>();
		try{
			workbook= new HSSFWorkbook(file.getInputStream());
	        Sheet sheet = workbook.getSheetAt(0);
	        Iterator<Row> rows = sheet.iterator();

	        while (rows.hasNext()) {
		       	Row currentRow = rows.next();
		       	Iterator<Cell> cellsInRow = currentRow.iterator();
		       	while (cellsInRow.hasNext()) {
		       		Cell currentCell = cellsInRow.next();
		       		list.add(currentCell.getStringCellValue());
		       	}
		       	break;
	        }
		}catch(Exception e){
			list=null;
			e.printStackTrace();
		}finally{
			 if(workbook!=null){
				 workbook.close();
			 }
		}
		return list;
	}

	public String uploadExcel(HashMap<String, String> map, MultipartFile file) throws Exception {
    	String result="success";

    	try{
    		init();
    		setTableInfo(map);
    		setColInfo(map);

			executeDDL();

			executeDML(file);

    		insertLayerInfo(map);

    	}catch(Exception e){
    		result=errorMsg;
    		if("".equals(result)){
    			result="서버에 에러가 발생했습니다.";
    		}
    		//에러 발생 했을 때, 생성된 테이블 drop, gis_lyr_list에서 생성된테이블 row 삭제
    		dropTable();
    		deleteLayerInfo();
    		e.printStackTrace();
    	}

    	return result;

    }


	private void setTableInfo(HashMap<String, String> map) throws UnsupportedEncodingException{
		this.tblNm=map.get("tblNm");
		this.tblKrNm=map.get("tblKrNm");
		this.lyrTyp=map.get("lyrTyp");
    }

	private void setColInfo(HashMap<String, String> map) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, UnsupportedEncodingException {
		this.colCnt=Integer.parseInt(map.get("colCnt"));

		for(int i=2; i<colCnt; i++){
			this.refColNm.invokeSetMethod(colNmObj, "colNm"+i, map.get("colNm"+i));
			this.refColKrNm.invokeSetMethod(colKrNmObj, "colKrNm"+i, map.get("colKrNm"+i));
			this.refColType.invokeSetMethod(colTypeObj, "colType"+i, map.get("colType"+i));
			this.refColLen.invokeSetMethod(colLenObj, "colLen"+i, map.get("colLen"+i));
		}
    }
//한글 인코딩이 안되어있어서 만든 함수. web.xml에서 인코딩 설정하니까 한글 인코딩이 됨. 그래서 주석 처리함.
//	private String changeCharset(String str) throws UnsupportedEncodingException {
//		String result = new String(StrUtil.chkNull(str).getBytes("8859_1"),"utf-8");
//		return result;
//    }

	private void executeDDL() throws SQLException, IllegalAccessException, IllegalArgumentException, InvocationTargetException{

		String tableName=tblNm;
		String tableKoreaName=tblKrNm;
		String layerType=replaceLyrTypToGeoType(lyrTyp);

		HashMap<String,Object> columnKrNmMap=makeColumnKrNmMap();
		HashMap<String,Object> columnTypeMap=makeColumnTypeMap();

		ExcelVo vo=new ExcelVo();

		vo.setTableName(tableName);
		vo.setTableKoreaName(tableKoreaName);
		vo.setLayerType(layerType);
		vo.setColumnKrNmMap(columnKrNmMap);
		vo.setColumnTypeMap(columnTypeMap);

		try{
			mapper.create(vo);
		}catch(Exception e){
			errorMsg="DDL에러";
			throw e;
		}
    }

	private String replaceLyrTypToGeoType(String str) {
		String result="";
		switch(str){
			case "point":   result="Point";
						    break;
			case "line":	result="MultiLineString";
							break;
			case "polygon": result="MultiPolygon";
							break;
		}
		return result;
	}

	private HashMap<String, Object> makeColumnKrNmMap() throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		String colNm;
		String colKrNm;

		HashMap<String,Object> columnKrNmMap=new HashMap<String,Object>();
		for(int i=2; i<colCnt; i++){
			colNm=refColNm.invokeGetMethod(colNmObj, "colNm"+i);
//			colType=refColType.invokeGetMethod(colTypeObj, "colType"+i);
			colKrNm=refColKrNm.invokeGetMethod(colKrNmObj, "colKrNm"+i);
			//value값이 숫자이면 string->double
			columnKrNmMap.put(colNm, colKrNm);
		}
		return columnKrNmMap;
	}

	private HashMap<String, Object> makeColumnTypeMap() throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		String colNm;
		String colType;
		String colLen;

		HashMap<String,Object> columnTypeMap=new HashMap<String,Object>();
		for(int i=2; i<colCnt; i++){
			colNm=refColNm.invokeGetMethod(colNmObj, "colNm"+i);
			colType=refColType.invokeGetMethod(colTypeObj, "colType"+i);
			colLen=refColLen.invokeGetMethod(colLenObj, "colLen"+i);

			columnTypeMap.put(colNm, addColTypeAndColLen(colType,colLen));
		}
		return columnTypeMap;
	}
	private String addColTypeAndColLen(String colType, String colLen) {
		String result=colType;

		if("text".equals(colType)){

		}else{
			result+="("+colLen+")";
		}
		return result;
	}

	private void executeDML(MultipartFile file) throws Exception {
		String[] splitList=file.getOriginalFilename().split(Pattern.quote("."));
		String excelType=splitList[splitList.length-1];

        if("xlsx".equals(excelType)){
        	if(!executeDMLAtXlsx(file)){
        		makeErrorMsg();
        		throw new Exception();
        	}
        }
        else if("xls".equals(excelType)){
        	if(!executeDMLAtXls(file)){
        		makeErrorMsg();
        		throw new Exception();
        	}
        }else{
        	errorMsg="엑셀파일은 xls,xlsx만 가능합니다.";
        	throw new Exception();
        }
	}

	private boolean executeDMLAtXlsx(MultipartFile file) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, IOException, SQLException {

		Workbook workbook=null;
		boolean result=true;

		try{
			HashMap<String, String> rowMap;

			workbook = new XSSFWorkbook(file.getInputStream());
	        Sheet sheet = workbook.getSheetAt(0);
	        Iterator<Row> rows = sheet.iterator();

	        Row currentRow;
	        int rowNumber=1;

			while (rows.hasNext()) {
				try{
					//첫번째 행은 컬럼명이어서 continue
					if (rowNumber ==1) {
						 currentRow = rows.next();
					     rowNumber++;
					     continue;
					}
					System.out.println("--------------------"+rowNumber);
					currentRow = rows.next();
					rowMap = makeRowMap(currentRow);
					//DB에  INSERT
					insertExcelRow(rowMap);
					rowNumber++;
				}catch(Exception e){
					errorMsg+=rowNumber+",";
					rowNumber++;
					result=false;
					e.printStackTrace();
				}
		    }

		}catch(Exception e){
			throw e;
		}finally{
			if(workbook!=null){
				workbook.close();
			}
		}
		return result;
	}

	private boolean executeDMLAtXls(MultipartFile file) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, IOException, SQLException {
		HSSFWorkbook workbook=null;
		boolean result=true;
		try{
			HashMap<String, String> rowMap;

			workbook = new HSSFWorkbook(file.getInputStream());
	        Sheet sheet = workbook.getSheetAt(0);
	        Iterator<Row> rows = sheet.iterator();

	        Row currentRow;
		    int rowNumber = 1;

			while (rows.hasNext()) {
				try{
					//첫번째 행은 컬럼명이어서 continue
					if (rowNumber ==1) {
						 currentRow = rows.next();
					     rowNumber++;
					     continue;
					}
					System.out.println("--------------------"+rowNumber);
					currentRow = rows.next();
					rowMap = makeRowMap(currentRow);
					//DB에  INSERT
					insertExcelRow(rowMap);
					rowNumber++;
				}catch(Exception e){
					errorMsg+=rowNumber+",";
					rowNumber++;
					result=false;
					e.printStackTrace();
				}
		    }
		}catch(Exception e){
			throw e;
		}finally{
			if(workbook!=null){
				workbook.close();
			}
		}
		return result;

	}

	private HashMap<String, String> makeRowMap(Row currentRow) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		String cellVal="";
		String key="";
		Cell currentCell;

		HashMap<String, String> rowMap=new HashMap<String, String>();

		int cnt=0;
		Iterator<Cell> cellsInRow= currentRow.iterator();

		while (cellsInRow.hasNext()) {
			currentCell = cellsInRow.next();
			//key
			if(cnt==0){
				key="_annox";
			}
			else if(cnt==1){
				key="_annoy";
			}
			else{
				key=refColNm.invokeGetMethod(colNmObj, "colNm"+cnt);
			}
			//value
			if(currentCell.getCellType()==0)//double이면
				cellVal=Double.toString(currentCell.getNumericCellValue());
			else //string이면
				cellVal=currentCell.getStringCellValue();

			rowMap.put(key, cellVal);
//			System.out.println(currentCell.getStringCellValue());
			cnt++;
		}
		return rowMap;
	}

	private void insertExcelRow(HashMap<String, String> rowMap) throws SQLException, IllegalAccessException, IllegalArgumentException, InvocationTargetException{

		if("point".equals(lyrTyp)){
			insertPoint(rowMap);
		}
		else if("line".equals(lyrTyp)){
			insertLine(rowMap);
		}
		else{
			insertPolygon(rowMap);
		}
    }

	private void insertPoint(HashMap<String, String> rowMap) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, SQLException {
		String tableName=tblNm;
		String _annox=rowMap.get("_annox").trim();
		String _annoy=rowMap.get("_annoy").trim();
		HashMap<String,Object> columnValueMap=makecolumnValueMap(rowMap);

		ExcelVo vo=new ExcelVo();

		vo.setTableName(tableName);
		vo.set_annox(_annox);
		vo.set_annoy(_annoy);
		vo.setPoint(makePointString(rowMap));
		vo.setColumnValueMap(columnValueMap);

		try{
			mapper.insertPoint(vo);
		}catch(Exception e){
			throw e;
		}
	}

	private String makePointString(HashMap<String, String> rowMap) {
		String _annox=rowMap.get("_annox").trim();
		String _annoy=rowMap.get("_annoy").trim();

		return "POINT("+_annox+" "+_annoy+")";
	}

	private void insertLine(HashMap<String, String> rowMap) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, SQLException {

		String tableName=tblNm;
		String _annox=rowMap.get("_annox").trim();
		String _annoy=rowMap.get("_annoy").trim();
		HashMap<String,Object> columnValueMap=makecolumnValueMap(rowMap);

		ExcelVo vo=new ExcelVo();

		vo.setTableName(tableName);
		vo.set_annox(_annox);
		vo.set_annoy(_annoy);
		vo.setLine(makeLineString(rowMap));
		vo.setColumnValueMap(columnValueMap);

		try{
			mapper.insertLine(vo);
		}catch(Exception e){
			throw e;
		}
	}

	private String makeLineString(HashMap<String, String> rowMap) {
		String annox[]=rowMap.get("_annox").split(",");
		String annoy[]=rowMap.get("_annoy").split(",");

		String str="MULTILINESTRING((";
		for(int i=0; i<annox.length; i++){
			str+=annox[i].trim()+" "+annoy[i].trim();
			str+=",";
		}
		str=str.substring(0,str.length()-1);
		str+="))";

		return str;
	}
	private void insertPolygon(HashMap<String, String> rowMap) {
		// TODO Auto-generated method stub

	}

	private HashMap<String, Object> makecolumnValueMap(HashMap<String, String> rowMap) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		String colNm;
		String value;
		String colType;

		HashMap<String,Object> columnValueMap=new HashMap<String,Object>();
		for(int i=2; i<colCnt; i++){
			colNm=refColNm.invokeGetMethod(colNmObj, "colNm"+i);
			colType=refColType.invokeGetMethod(colTypeObj, "colType"+i);
			value=rowMap.get(colNm);
			//value값이 숫자이면 string->double
			if("numeric".equals(colType)){
				columnValueMap.put(colNm, Double.parseDouble(value));
			}
			else{
				columnValueMap.put(colNm, value);
			}
		}
		return columnValueMap;
	}

	private void makeErrorMsg() {
		errorMsg=errorMsg.substring(0,errorMsg.length()-1);
		errorMsg="엑셀 파일 "+errorMsg+"에 문제가 있습니다. 확인해주시길 바랍니다.";
	}

	private void insertLayerInfo(HashMap<String, String> map) throws Exception {
		LayerVo vo=makeLayerVo();
		LayerStyleVo styleVo=new LayerStyleVo();
		//gis_lyr_list에  insert
		LayerVo affectVo=svc.addAndgetItem(vo);
		//gis_lyr_style에  insert
		styleVo.setLyrMgrSeq(affectVo.getMgrSeq());
		style.add(styleVo);
	}

	private LayerVo makeLayerVo() throws Exception {
		LayerVo vo=new LayerVo();
		vo.setSchemNm(EXCEL_SCHEMA);
		vo.setTblId(tblNm);
		vo.setLyrNm(tblKrNm);
		vo.setLyrTyp(replaceLyrTypToChar(lyrTyp));
		vo.setGrpMgrSeq(this.getGrpMgrSeql());
		vo.setMkUser("user");//추후에 session으로 가져와야함.
		return vo;
	}

	private int getGrpMgrSeql() throws Exception {
		HashMap<String, String> map=new HashMap<String, String>();
		map.put("mkUser","user"); //추후에 session으로 가져와야함.
		map.put("grpNm","나의레이어");
		LayerGroupVo groupVO=group.getItem(map);
		return groupVO.getMgrSeq();
	}

	private String replaceLyrTypToChar(String str) {
		String result="";
		switch(str){
			case "point":   result="P";
						    break;
			case "line":	result="L";
							break;
			case "polygon": result="G";
							break;
		}
		return result;
	}

	private void dropTable() {
		String tableName=tblNm;
		ExcelVo vo=new ExcelVo();
		vo.setTableName(tableName);
		try{
			mapper.drop(vo);
		}catch(Exception e){
			throw e;
		}
	}
	private void deleteLayerInfo() {
		LayerVo vo=new LayerVo();
		vo.setSchemNm(EXCEL_SCHEMA);
		vo.setTblId(tblNm);
		svc.delExcelLayer(vo);

	}

}