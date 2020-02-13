package com.example.demo1.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.util.StringUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.List;
import java.util.UUID;

/**
 * @author haozz
 * @date 2018/6/6 9:57
 * @description excel导出抽象工具类
 **/
public abstract class ExportAbstractUtil {
    public void write(HttpServletResponse response, Workbook workbook){
        String fileName = UUID.randomUUID().toString()+".xlsx";
        pwrite(response,workbook,fileName);
    }
    public void write(HttpServletResponse response,Workbook workbook,String fileName){
        if(StringUtils.isEmpty(fileName)){
            fileName = UUID.randomUUID().toString()+".xlsx";
        }
        pwrite(response,workbook,fileName);
    }
    public void write(HttpServletResponse response, List<List<String>> lists,String fileName){
        if(StringUtils.isEmpty(fileName)){
            fileName = UUID.randomUUID().toString()+".xlsx";
        }
        SXSSFWorkbook workbook = new SXSSFWorkbook(lists.size());
        SXSSFSheet sheet = workbook.createSheet(fileName.substring(0,fileName.indexOf(".xls")));
        Integer rowIndex = 0;
        Row row = null;
        Cell cell = null;
        for(List<String> rowData: lists ){
            Integer columnIndex = 0;
            row = sheet.createRow(rowIndex++);
            for(String columnVal:rowData){
                cell = row.createCell(columnIndex++);
                cell.setCellValue(columnVal);
            }
        }
        pwrite(response,workbook,fileName);
    }
    private void pwrite(HttpServletResponse response,Workbook workbook,String fileName){
        response.setCharacterEncoding("UTF-8");
        response.setContentType("application/vnd.ms-excel;charset=UTF-8");
        try {
            response.addHeader("Content-Disposition", "attachment; filename="+new String(fileName.getBytes("UTF-8"),"ISO8859-1"));
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
            fileName= UUID.randomUUID().toString()+".xlsx";
            response.addHeader("Content-Disposition", "attachment; filename="+fileName);
        }
        try {
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}