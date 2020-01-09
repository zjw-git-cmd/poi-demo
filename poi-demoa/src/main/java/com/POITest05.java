package com;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author zhengjiawei
 * @date 2020/1/9 15:26
 * 读取Excel表格,并解析内容
 */

public class POITest05 {
    public static void main(String[] args) throws IOException {
        //1.根据Excel表格创建工作簿
        Workbook wb=new XSSFWorkbook("D:\\poi\\zjw03.xlsx");
        //2.获取Sheet   根据索引,获取对应的Sheet
        Sheet sheet = wb.getSheetAt(0);
        //3.获取Sheet中的每一行,和每一个单元格
        for(int rowNum=0;rowNum<sheet.getLastRowNum();rowNum++){
            Row row = sheet.getRow(rowNum);
            StringBuilder sb=new StringBuilder();
            for(int cellNum=0;cellNum<row.getLastCellNum();cellNum++){
                //根据索引获取每一个单元格
                Cell cell = row.getCell(cellNum);
                //获取每一个单元格的内容


            }
            System.out.println(sb.toString());

        }
    }
    public static Object getCellValue(Cell cell){
        //获取单元格的属性类型
        CellType cellType = cell.getCellType();
        Object value=null;
        //根据单元格数据类型获取数据
        switch (cellType){
            case STRING:
                value=cell.getStringCellValue();
                break;
            case BOOLEAN:
                value=cell.getBooleanCellValue();
                break;
            case NUMERIC:
                //判断是否是日期格式
                if(DateUtil.isCellDateFormatted(cell)){
                    //日期格式
                    value=cell.getDateCellValue();
                }else{
                    //是数字
                    value=cell.getNumericCellValue();
                }
            case FORMULA:
        }
        return null;
    }
}
