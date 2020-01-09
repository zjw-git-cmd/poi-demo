package com;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author zhengjiawei
 * @date 2020/1/9 15:26
 * 在指定的文件夹下创建了一个空的表格,注意是2007版本的,文件的后缀名称是xlsx
 */

public class POITest01 {
    public static void main(String[] args) throws IOException {
        //1.创建一个工作簿  HSSFWorkbook--->创建的是2003版本的表格
        Workbook wb=new XSSFWorkbook();//2007版本的表格
        //2.创建表单sheet
        Sheet sheet=wb.createSheet("test");
        //3.文件流
        FileOutputStream fos=new FileOutputStream("D:\\poi\\zjw.xlsx");
        //4.写入文件
        wb.write(fos);
        //5.关闭文件输出流
        fos.close();
    }
}
