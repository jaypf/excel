package com.example.excel.deal;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.Data;

/**
 * @ClassName TestModel
 * @Description TODO
 * @Author Jay.Jia
 * @Date 2020/6/29 11:25
 * @Version 1.0
 */
@Data
public class TestModel extends BaseRowModel {

        @ExcelProperty(value = "姓名", index = 0)
        private String name;
        @ExcelProperty(value = "微信号", index = 1)
        private String weChat;
        @ExcelProperty(value = "手机号", index = 2)
        private String phone;

}
