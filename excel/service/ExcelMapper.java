package gmx.gis.excel.service;


import egovframework.rte.psl.dataaccess.mapper.Mapper;

/**
 * @author 민동현
 *
 */
@Mapper("excelMapper")
public interface ExcelMapper {

	void create(ExcelVo vo);

	void drop(ExcelVo vo);

	void insertPoint(ExcelVo vo);

	void insertLine(ExcelVo vo);



}
