package com.tsn.util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * Excel工具类
 * @author zhaoyuanyuan
 */
public class ExcelUtil {

    /**
     * 导出Excel
     *
     * @param sheetName sheet名称
     * @param title     标题
     * @param values    内容
     * @param wb        HSSFWorkbook对象
     * @return
     */
    public static HSSFWorkbook getHSSFWorkbook(String sheetName, String[] title, String[][] values, HSSFWorkbook wb) {

        // 第一步，创建一个HSSFWorkbook，对应一个Excel文件
        if (wb == null)
            wb = new HSSFWorkbook();

        // 第二步，在workbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet sheet = wb.createSheet(sheetName);

        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制
        HSSFRow row = sheet.createRow(0);

        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = wb.createCellStyle();
        //设置水平垂直居中
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER); //水平
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直
        //设置自动换行
      	style.setWrapText(true);
      	//设置边框
      	style.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框  
      	style.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框  
      	style.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框  
      	style.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
      	
        //声明列对象
        HSSFCell cell = null;
        
        //设置列宽
        sheet.setColumnWidth(0, 1500);
        sheet.setColumnWidth(1, 3800);
        sheet.setColumnWidth(2, 3000);
        sheet.setColumnWidth(3, 3000);
        sheet.setColumnWidth(4, 3000);
        sheet.setColumnWidth(5, 12000);
        sheet.setColumnWidth(6, 3800);
        sheet.setColumnWidth(7, 3800);
        
        //创建标题
        for (int i = 0; i < title.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(style);
        }

        //创建内容
        for (int i = 0; i < values.length; i++) {
            row = sheet.createRow(i + 1);
            for (int j = 0; j < values[i].length; j++) {
                //将内容按顺序赋给对应的列对象
            	cell =  row.createCell(j);
                cell.setCellValue(values[i][j]);
                //设置单元格内容居中
                cell.setCellStyle(style);
            }
        }
        
        //设置打印参数
        HSSFPrintSetup ps = sheet.getPrintSetup();
        ps.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)   
        ps.setPaperSize(HSSFPrintSetup.A4_PAPERSIZE); //纸张类型  
        sheet.setHorizontallyCenter(true);//设置打印页面为水平居中
        
        return wb;
    }
}