package com.bochao.util;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 * @auth 吴军杰
 * @Description 导出数据为excel工具类
 * @date 2021-7-23 15:10:18
 */
public class HSSFWorkbookUtil {

    /**
     * 将数据转换为表格数据保存
     * @author 吴军杰
     * @date 2021/07/19 14:50
     * */
    public static void getCells(String[] data, HSSFRow row, HSSFWorkbook workbook, HSSFCellStyle style) {
        for (int i = 0; i < data.length; i++) {
            //创建HSSFCell对象
            HSSFCell cell = row.createCell(i);
            //设置单元格的值
            cell.setCellValue(data[i]);
            // 设置样式
            cell.setCellStyle(style);
        }
    }

    /**
     * 设置表格单元格样式
     * @author 吴军杰
     * @date 2021/07/19 14:50
     * */
    public static HSSFCellStyle getCellStyle(HSSFWorkbook workbook) {
        // 设置样式
        HSSFCellStyle style = workbook.createCellStyle();
        //设置字体
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 14);
        style.setFont(font);
        // 设置居中
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }
}
