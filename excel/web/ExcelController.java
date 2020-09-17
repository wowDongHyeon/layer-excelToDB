package gmx.gis.excel.web;



import java.util.HashMap;

import javax.annotation.Resource;
import javax.servlet.http.HttpSession;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.validation.Validator;
import org.springframework.web.bind.WebDataBinder;
import org.springframework.web.bind.annotation.InitBinder;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.multipart.MultipartFile;

import gmx.gis.excel.service.ExcelService;
import gmx.gis.excel.util.ExcelUtil;

/**
 * @author 민동현
 *
 */
@Controller
@RequestMapping("/excel")
public class ExcelController {

    @Resource(name = "excelService")
    private ExcelService excelService;

    @Resource
    private Validator validator;


    @InitBinder
    private void initBinder(WebDataBinder binder){
        binder.setValidator(this.validator);
    }


    @RequestMapping(value = "/getExcelColumnList.json", method = RequestMethod.POST)
    public void getExcelColumnList(Model model, HttpSession session, @RequestPart("file") MultipartFile file) throws Exception {
         try{
        	 model.addAttribute("excelColumnList", excelService.getExcelColumnList(file));
         }catch(Exception e){
        	 model.addAttribute("error", "엑셀파일은 xls,xlsx만 가능합니다.");
         }

    }

    @RequestMapping(value = "/uploadExcel.json", method = RequestMethod.POST)
    public void uploadExcel(Model model, HttpSession session, @RequestParam HashMap<String, String> map,@RequestPart("file") MultipartFile file) throws Exception {
        map.put("userId", (String) session.getAttribute("userId"));
        String msg = ExcelUtil.findError(map);
        if("pass".equals(msg)){
            model.addAttribute("result", excelService.uploadExcel(map, file));
        }else{
            model.addAttribute("result", msg);
        }
    }

}
