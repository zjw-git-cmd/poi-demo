package com;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author zhengjiawei
 * @date 2020/1/9 15:26
 * 在指定的文件夹下创建了一个空的表格,注意是2007版本的,文件的后缀名称是xlsx
 * 创建单元格,并写入内容
 */

public class POITest02 {
    public static void main(String[] args) throws IOException {
        //创建一个工作簿  HSSFWorkbook--->创建的是2003版本的表格
        Workbook wb=new XSSFWorkbook();//2007版本的表格
        //创建表单sheet
        Sheet sheet = wb.createSheet("表单1");
        //创建行对象   参数:索引(从0开始)
        Row row = sheet.createRow(2);
        //创建单元格对象  参数:索引(从0开始)
        Cell cell=row.createCell(2);
        //向单元格中写入内容
        cell.setCellValue("这个是我创建的第一个单元格");
        //文件流
        FileOutputStream fos=new FileOutputStream("D:\\poi\\zjw01.xlsx");
        //4.写入文件
        wb.write(fos);
        //5.关闭文件输出流
        fos.close();
    }
}
