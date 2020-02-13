package com.example.demo1.controller;

import com.example.demo1.utils.ExportAbstractUtil;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * @author haozz
 * @date 2018/6/6 10:14
 * @description
 **/
@Controller
@RequestMapping(value = "/index")
public class HelloCtrl extends ExportAbstractUtil {
    @RequestMapping(value = "/testExcelOutPut")
    @ResponseBody
    public void testExcelOutPut(HttpServletResponse response){
        //拼接数据start
        List<List<String>> lists = new ArrayList<List<String>>();
        String rows[] = {"tian","month","day","报表"};
        List<String> rowsTitle = Arrays.asList(rows);
        lists.add(rowsTitle);
        for(int i = 0; i<=9;i++){
            String [] rowss = {"1","2","3"};
            List<String> rowssList = Arrays.asList(rowss);
            lists.add(rowssList);
        }
        //拼接数据end
        write(response,lists,"导出Excel.xlsx");
    }

}