package com;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author zhengjiawei
 * @date 2020/1/9 15:26
 * 在指定的文件夹下创建了一个空的表格,注意是2007版本的,文件的后缀名称是xlsx
 * 创建单元格,并写入内容
 * 设定单元格和内容的样式
 */

public class POITest03 {
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
        /**
         * 样式处理
         */
        //创建单元格样式对象
        CellStyle cellStyle=wb.createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);//上边框
        cellStyle.setBorderBottom(BorderStyle.THIN);//下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellStyle.setBorderRight(BorderStyle.THIN);//右边框
        //创建字体对象
        Font font=wb.createFont();
        font.setFontName("华文行楷");//字体样式
        font.setFontHeightInPoints((short)28);//字体大小
        //将字体对象,放入单元格对象中
        cellStyle.setFont(font);
        //设定行高和列宽
        row.setHeightInPoints(50);
        //注意源码中,将列的宽度进行了÷256    setColumnWidth(m,n):m代表的是索引,就是第几列,n代表的是列的宽度
        sheet.setColumnWidth(2,31*256);
        //设定单元格对齐方式
        cellStyle.setAlignment(HorizontalAlignment.CENTER);//水平居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
        //将样式放入单元格中
        cell.setCellStyle(cellStyle);
        //文件流
        FileOutputStream fos=new FileOutputStream("D:\\poi\\zjw01.xlsx");
        //4.写入文件
        wb.write(fos);
        //5.关闭文件输出流
        fos.close();
    }
}
