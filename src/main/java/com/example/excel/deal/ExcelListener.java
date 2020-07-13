package com.example.excel.deal;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.util.ArrayList;
import java.util.List;

/**
 * @ClassName ExcelListener
 * @Description TODO
 * @Author Jay.Jia
 * @Date 2020/6/29 11:26
 * @Version 1.0
 */
public class ExcelListener extends AnalysisEventListener {
    //自定义用于暂时存储data
    //private List<Object> datas = Collections.synchronizedList(new ArrayList<>());
    private List<Object> datas = new ArrayList<>();

    /**
     * 通过 AnalysisContext 对象还可以获取当前 sheet，当前行等数据
     */
    @Override
    public void invoke(Object o, AnalysisContext analysisContext) {
        datas.add(o);
    }

    /**
     * 读取完之后的操作
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {

    }

    public List<Object> getDatas() {
        return datas;
    }

    public void setDatas(List<Object> datas) {
        this.datas = datas;
    }
}
