package com.example.excel.deal;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * @ClassName DealController
 * @Description TODO
 * @Author Jay.Jia
 * @Date 2020/6/29 11:28
 * @Version 1.0
 */
@RestController
public class DealController {

    @Autowired
    private DealService service;

    //导入excel
    @RequestMapping(value = "excelImport", method = {RequestMethod.GET, RequestMethod.POST })
    public String excelImport(HttpServletRequest request, Model model, @RequestParam("uploadFile") MultipartFile[] files) throws Exception {
        if (files != null && files.length > 0) {
            MultipartFile file = files[0];
            List<Object> list = EasyExcelUtil.readExcel(file, new TestModel(), 1, 1);
            if (list != null && list.size() > 0) {
                for (Object o : list) {
                    TestModel xfxx = (TestModel) o;
                    System.out.println(xfxx.getName() + "/" + xfxx.getWeChat() + "/" + xfxx.getPhone());
                }
            }
        }
        return "index";
    }

    @GetMapping("index")
    public String index(){
        return "indexadasd";
    }

    @GetMapping("deaul")
    public String deaul(@RequestParam Map<String, String> parm, HttpServletResponse response){

        try {
//            service.deal(parm, response);
            service.deal2local(parm, response);
        } catch (Exception e) {
            e.printStackTrace();
            return "error"+e.getMessage();
        }
        return "ok!";
    }

}
