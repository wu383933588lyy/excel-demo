package com.demo.write;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;

public class ExcelWriteTest {

    /**
     * 导出 xls 文件
     *
     * @throws Exception
     */
    @Test
    public void testWrite03() throws Exception {
        // 1. 创建工作簿
        Workbook workbook = new HSSFWorkbook();
        // 2. 创建工作表
        Sheet sheet = workbook.createSheet();
        // 3. 创建行
        Row row0 = sheet.createRow(0);
        // 4. 创建单元格
        Cell cell = row0.createCell(0);
        cell.setCellValue("今日新增关注");
        cell = row0.createCell(1);
        cell.setCellValue("999");

        FileOutputStream stream = new FileOutputStream("F:\\CodeSpace\\excel-demo\\03.xls");
        workbook.write(stream);
        stream.close();
    }

    /**
     * 导出 xlsx 文件
     * 使用
     * @throws Exception
     */
    @Test
    public void testWrite07() throws Exception {
        // 1. 创建工作簿
        Workbook workbook = new XSSFWorkbook();
        // 2. 创建工作表
        Sheet sheet = workbook.createSheet();
        // 3. 创建行
        Row row0 = sheet.createRow(0);
        // 4. 创建单元格
        Cell cell = row0.createCell(0);
        cell.setCellValue("今日新增关注");
        cell = row0.createCell(1);
        cell.setCellValue("999");

        FileOutputStream stream = new FileOutputStream("F:\\CodeSpace\\excel-demo\\07.xlsx");
        workbook.write(stream);
        stream.close();
    }
}
