package com.example.excel.deal;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Builder;
import lombok.Data;

/**
 * @ClassName ResVO
 * @Description TODO
 * @Author Jay.Jia
 * @Date 2020/6/29 15:55
 * @Version 1.0
 */
@Data
@Builder
public class ResVO extends TestModel{
    @ExcelProperty(value = "文件名", index = 3)
    private String fileName;
    @ExcelProperty(value = "行数", index = 4)
    private int rowNum;

}
