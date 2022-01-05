package com.bochao.file;

import com.bochao.entity.DeviceInfo;
import com.bochao.util.HSSFWorkbookUtil;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.springframework.util.StringUtils;
import java.io.*;
import java.nio.charset.Charset;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipInputStream;

/**
 * 模型文件数据转换
 * */
@Log4j2
public class FileParse {
    public static void main(String[] args) {
        // 每行信息
        Row row = null;
        try {
            log.info("开始解析编码规范文件");
            // 一. 解析编码规范文件
            // 1- 读取编码规范文件
            FileInputStream fileInputStream = new FileInputStream("C:/file/编码规范.xlsx");
            // 2- 获取 wookbook
            Workbook workbook = WorkbookFactory.create(fileInputStream);
            // 3- 获取sheet
            Sheet sheet = workbook.getSheetAt(0);
            // 4- 获取总行数
            int rowLength = sheet.getLastRowNum();
            // 5- 从第二行开始读取数据
            // 设备信息数据
            DeviceInfo deviceInfo = new DeviceInfo();
            // 全部设备信息map
            Map<String, DeviceInfo> map = new HashMap<>();
            for (int i = 1; i < rowLength; i++){
                // 5.1 获取当前行
                row = sheet.getRow(i);
                // 5.2 获取每列数据
                if (row != null) {
                    DeviceInfo newDeviceInfo = new DeviceInfo();
                    // 区域
                    newDeviceInfo.setArea(getColumnValue(row, 0, deviceInfo.getArea()));
                    // 设备名称
                    newDeviceInfo.setDeviceName(getColumnValue(row, 1, deviceInfo.getDeviceName()));
                    // 部件
                    newDeviceInfo.setParts(getColumnValue(row, 2, deviceInfo.getParts()));
                    // 编码
                    newDeviceInfo.setCode(getColumnValue(row, 3, deviceInfo.getCode()));
                    // 模型名称
                    String modelName = getColumnValue(row, 4, deviceInfo.getModelName());
                    newDeviceInfo.setModelName(modelName);
                    map.put(modelName, newDeviceInfo);
                    deviceInfo = newDeviceInfo;
                }
            }
            fileInputStream.close();
            log.info("解析编码规范文件结束");
            // 二. 夺取压缩文件模型名称数据
            // 1- 获取压缩文件
            String path = "C:/file/荣信阀塔.zip";
            ZipFile zipFile = new ZipFile(path, Charset.forName("GBK"));
            InputStream in = new BufferedInputStream(new FileInputStream(path));
            ZipInputStream zin = new ZipInputStream(in, Charset.forName("GBK"));
            ZipEntry ze;
            // 保存压缩文件内容list
            List<String> modelNameList = new ArrayList<>();
            while ((ze = zin.getNextEntry()) != null) {
                if (!ze.isDirectory()) {
                    long size = ze.getSize();
                    if (size > 0) {
                        BufferedReader br = new BufferedReader(new InputStreamReader(zipFile.getInputStream(ze), Charset.forName("gbk")));
                        String line;
                        while ((line = br.readLine()) != null) {
                            String[] index = line.split(",");
                            modelNameList.add(index[0]);
                        }
                        br.close();
                    }
                }
            }
            zin.closeEntry();
            // 三. 导出编码信息数据
            // 1- 创建一个工作表sheet
            HSSFWorkbook writeWorkbook = new HSSFWorkbook();
            // 2 创建HSSFSheet对象
            HSSFSheet writeSheet = writeWorkbook.createSheet("sheet0");
            // 3 设置默认行宽
            writeSheet.setDefaultColumnWidth(20);
            // 4 设置单元格样式
            HSSFCellStyle cellStyle = HSSFWorkbookUtil.getCellStyle(writeWorkbook);
            // 5 表头信息处理
            // 创建HSSFRow对象，创建表头信息行
            HSSFRow wirteRow = writeSheet.createRow(0);
            // 表头信息
            String[] data = {"区域", "设备", "部件", "编码", "模型名称"};
            // 将表头信息转换为表格数据保存
            HSSFWorkbookUtil.getCells(data, wirteRow, writeWorkbook, cellStyle);
            // 6 将模型数据转换为表格数据导出
            Integer num = 0;
            for (String modelName : modelNameList) {
                // 从map中获取对应的设备数据
                // 获取最后一个_前的数据
                int index = modelName.lastIndexOf("_");
                String leftName  = modelName.substring(0, index);
                String rightName  = modelName.substring(index,  modelName.length());
                DeviceInfo writeInfo = map.get(leftName);
                if (writeInfo != null) {
                    num = num + 1;
                    wirteRow = writeSheet.createRow(num);
                    String[] measureData = {writeInfo.getArea(), writeInfo.getDeviceName(), writeInfo.getParts(),writeInfo.getCode() + rightName, writeInfo.getModelName() + rightName};
                    HSSFWorkbookUtil.getCells(measureData, wirteRow, writeWorkbook, cellStyle);
                }
            }
            // 5 输出Excel文件
            // 打开文件流
            File file = new File("C:file/编码规范" + new Date().getTime() + ".csv");
            file.createNewFile();
            FileOutputStream outputStream = new FileOutputStream(file);
            writeWorkbook.write(outputStream);
            outputStream.close();
        } catch (Exception e) {
            log.info("失败" + row);
            e.printStackTrace();
        }
    }

    // 解析编码规范文件

    // 获取列的值
    public static String getColumnValue(Row row, int i, String value) {
        Cell cell = row.getCell(i);
        if (cell != null) {
            String stringCellValue = cell.getStringCellValue();
            if (!StringUtils.isEmpty(stringCellValue)) {
                return stringCellValue;
            }
        }
        return value;
    }
}
