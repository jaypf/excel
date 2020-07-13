package com.example.excel.deal;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @ClassName EasyExcelUtil
 * @Description TODO
 * @Author Jay.Jia
 * @Date 2020/6/29 11:27
 * @Version 1.0
 */
public class EasyExcelUtil {

    /**
     * 读取某个 sheet 的 Excel
     *
     * @param excel    文件
     * @param rowModel 实体类映射，继承 BaseRowModel 类
     *  sheetNo  sheet 的序号 从1开始
     * @return Excel 数据 list
     */
    public static List<Object> readExcel(MultipartFile excel, BaseRowModel rowModel) throws IOException {
        return readExcel(excel, rowModel, 1, 1);
    }

    /**
     * 读取某个 sheet 的 Excel
     * @param excel       文件
     * @param rowModel    实体类映射，继承 BaseRowModel 类
     * @param sheetNo     sheet 的序号 从1开始
     * @param headLineNum 表头行数，默认为1
     * @return Excel 数据 list
     */
    public static List<Object> readExcel(MultipartFile excel, BaseRowModel rowModel, int sheetNo, int headLineNum) throws IOException {
        ExcelListener excelListener = new ExcelListener();
        ExcelReader reader = getReader(excel, excelListener);
        if (reader == null) {
            return null;
        }
        reader.read(new Sheet(sheetNo, headLineNum, rowModel.getClass()));
        return excelListener.getDatas();
    }

    /**
     * 读取指定sheetName的Excel(多个 sheet)
     * @param excel    文件
     * @param rowModel 实体类映射，继承 BaseRowModel 类
     * @return Excel 数据 list
     * @throws IOException
     */
    public static List<Object> readExcel(MultipartFile excel, BaseRowModel rowModel,String sheetName) throws IOException {
        ExcelListener excelListener = new ExcelListener();
        ExcelReader reader = getReader(excel, excelListener);
        if (reader == null) {
            return null;
        }
        for (Sheet sheet : reader.getSheets()) {
            if (rowModel != null) {
                sheet.setClazz(rowModel.getClass());
            }
            //读取指定名称的sheet
            if(sheet.getSheetName().contains(sheetName)){
                reader.read(sheet);
                break;
            }
        }
        return excelListener.getDatas();
    }

    public static List<Object> readExcelByPath(String filename, String path, BaseRowModel rowModel) throws IOException {
        ExcelListener excelListener = new ExcelListener();
        ExcelReader reader = getReader(filename, path, excelListener);
        if (reader == null) {
            return null;
        }

        for (Sheet sheet : reader.getSheets()) {
            if (rowModel != null) {
                sheet.setClazz(rowModel.getClass());
            }
            //读取sheet
            reader.read(sheet);
        }
        return excelListener.getDatas();
    }

    public static Map<String, List<Object>> readExcelByPath2Map(String filename, String path, BaseRowModel rowModel) throws IOException {
        //List<Object> list = new ArrayList<>();
//        Map<String, Map<String, List<Object>>> fileMap = new HashMap<>(16);
        Map<String, List<Object>> sheetMap = new HashMap<>(16);
//        fileMap.put(filename, sheetMap);

        ExcelListener excelListener = new ExcelListener();
        ExcelReader reader = getReader(filename, path, excelListener);
        if (reader == null) {
            return null;
        }

        for (Sheet sheet : reader.getSheets()) {
            if (rowModel != null) {
                sheet.setClazz(rowModel.getClass());
            }
            //读取sheet
            reader.read(sheet);
        }
        sheetMap.put(filename, excelListener.getDatas());
        return sheetMap;
    }

    /**
     * 返回 ExcelReader
     * @param excel 需要解析的 Excel 文件
     * @param excelListener new ExcelListener()
     * @throws IOException
     */
    private static ExcelReader getReader(MultipartFile excel, ExcelListener excelListener) throws IOException {
        String filename = excel.getOriginalFilename();
        if(filename != null && (filename.toLowerCase().endsWith(".xls") || filename.toLowerCase().endsWith(".xlsx"))){
            InputStream is = new BufferedInputStream(excel.getInputStream());
            return new ExcelReader(is, null, excelListener, false);
        }else{
            return null;
        }
    }

    private static ExcelReader getReader(String filename, String path, ExcelListener excelListener) throws IOException {
        if(filename != null && (filename.toLowerCase().endsWith(".xls") || filename.toLowerCase().endsWith(".xlsx"))){
            InputStream is = new BufferedInputStream(new FileInputStream(path));
            return new ExcelReader(is, null, excelListener, false);
        }else{
            return null;
        }
    }

    // 简单读取 (同步读取)
    public static void simpleRead() {

        // 读取 excel 表格的路径
        String readPath = "C:\\Users\\oukele\\Desktop\\模拟数据.xlsx";

        try {
            // sheetNo --> 读取哪一个 表单
            // headLineMun --> 从哪一行开始读取( 不包括定义的这一行，比如 headLineMun为2 ，那么取出来的数据是从 第三行的数据开始读取 )
            // clazz --> 将读取的数据，转化成对应的实体，需要 extends BaseRowModel
            Sheet sheet = new Sheet(1, 1, TestModel.class);

            // 这里 取出来的是 ExcelModel实体 的集合
            List<Object> readList = EasyExcelFactory.read(new FileInputStream(readPath), sheet);
            // 存 ExcelMode 实体的 集合
            List<TestModel> list = new ArrayList<>();
            for (Object obj : readList) {
                list.add((TestModel) obj);
            }

            // 取出数据
            StringBuilder str = new StringBuilder();
            str.append("{");
            String link = "";
            for (TestModel mode : list) {
                str.append(link).append("\""+mode.getName()+"\":").append("\""+mode.getPhone()+"\"");
                link= ",";
            }
            str.append("};");
            System.out.println(str);

        } catch (Exception e) {
            e.printStackTrace();
        }


    }

}
