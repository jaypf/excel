package com.example.excel.deal;

import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * @ClassName DealService
 * @Description TODO
 * @Author Jay.Jia
 * @Date 2020/6/29 15:33
 * @Version 1.0
 */
@Service
public class DealService {

    @Value("${filePath}")
    private String filePath;
    @Value("${resFilePath}")
    private String resFilePath;

    public void deal(Map<String, String> parm, HttpServletResponse response) throws Exception{

        ServletOutputStream out = response.getOutputStream();
        ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX, true);
        String fileName = new String((new SimpleDateFormat("yyyy-MM-dd").format(new Date()))
                .getBytes(), "UTF-8");
        Sheet sheet1 = new Sheet(1, 0, ResVO.class);
        sheet1.setSheetName("第一个sheet");
        writer.write(getListString(parm), sheet1);
        writer.finish();
        response.setContentType("multipart/form-data");
        response.setCharacterEncoding("utf-8");
        response.setHeader("Content-disposition", "attachment;filename="+fileName+".xlsx");
        out.flush();
    }

    public void deal2local(Map<String, String> parm, HttpServletResponse response) throws Exception{
        OutputStream out = new FileOutputStream(resFilePath);
//        ServletOutputStream out = response.getOutputStream();
        ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX, true);
        String fileName = new String((new SimpleDateFormat("yyyy-MM-dd").format(new Date()))
                .getBytes(), "UTF-8");
        Sheet sheet1 = new Sheet(1, 0, ResVO.class);
        sheet1.setSheetName("第一个sheet");
        writer.write(getListString(parm), sheet1);
        writer.finish();
        out.flush();
        out.close();
    }

    private List<ResVO> getListString(Map<String, String> parm) throws Exception{
//        String filePath = (String)request.getAttribute("path");
        String phone = parm.get("phone");
        List<ResVO> list = new ArrayList<>();
        File file = new File(filePath);
        if (file.isDirectory()) {
            System.out.println("文件夹");
            String[] filelist = file.list();
            for (int i = 0; i < filelist.length; i++) {
                File readfile = new File(filePath + "\\" + filelist[i]);
                if (!readfile.isDirectory()) {
                    String path = readfile.getPath();
                    String name = readfile.getName();
                    Map<String, List<Object>> stringListMap = EasyExcelUtil.readExcelByPath2Map(name, path, new TestModel());
                    stringListMap.forEach((k,v) ->{
                        v.forEach(t -> {
                            TestModel m = ((TestModel)t);
                            if (m.getPhone().startsWith(phone)){
                                ResVO build = ResVO.builder().fileName(k).rowNum(v.indexOf(t)+2).build();
                                BeanUtils.copyProperties(m, build);
                                list.add(build);
                            }
                        });
                    });
                }
            }
        }
        return list;
    }

}
