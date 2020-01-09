package com;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

/**
 * @author zhengjiawei
 * @date 2020/1/9 15:26
 * 向单元格中插入图片
 */

public class POITest04 {
    public static void main(String[] args) throws IOException {
        //创建一个工作簿  HSSFWorkbook--->创建的是2003版本的表格
        Workbook wb=new XSSFWorkbook();//2007版本的表格
        //创建表单sheet
        Sheet sheet = wb.createSheet("表单2");
        //读取图片的数据流
        FileInputStream fis=new FileInputStream("D:\\poi\\1.jpg");
        //转化为二进制数组
        byte[] bytes= IOUtils.toByteArray(fis);
        fis.read(bytes);
        //向POI内存中添加一张图片,返回图片在图片集合中的索引
        int index=wb.addPicture(bytes,Workbook.PICTURE_TYPE_JPEG);//参数一:图片的二进制数据,参数二:图片的类型
        //绘制图片的工具类
        CreationHelper helper=wb.getCreationHelper();
        //创建一个绘制图片对象
        Drawing<?> patriarch=sheet.createDrawingPatriarch();
        //创建锚点,设定图片的坐标
        ClientAnchor anchor=helper.createClientAnchor();
        anchor.setRow1(0);
        anchor.setCol1(0);
        //绘制图片
        Picture picture = patriarch.createPicture(anchor, index);
        picture.resize();//自使用渲染图片
        //文件流
        FileOutputStream fos=new FileOutputStream("D:\\poi\\zjw02.xlsx");
        //4.写入文件
        wb.write(fos);
        //5.关闭文件输出流
        fos.close();
    }
}
