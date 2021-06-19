package com.demo.write;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;

public class BigExcelWrite {

    /**
     *
     * @throws Exception
     */
    @Test
    public void bigExcel03()throws Exception{
        long start = System.currentTimeMillis();
        // 1. 创建工作簿
        Workbook workbook = new HSSFWorkbook();
        // 2. 创建工作表
        Sheet sheet = workbook.createSheet();
        // 3. 创建行
        Row row = null;
        Cell cell;
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 20; cellNum++) {
                cell = row.createCell(cellNum);
                cell.setCellValue(String.format("%s行%s列",rowNum,cellNum));
            }
        }
        FileOutputStream stream = new FileOutputStream("F:\\CodeSpace\\excel-demo\\03.xls");
        workbook.write(stream);
        stream.close();
        System.out.println(System.currentTimeMillis() - start);
    }

    /**
     * 07 耗时长于 03
     *
     * @throws Exception
     */
    @Test
    public void bigExcel07()throws Exception{
        long start = System.currentTimeMillis();
        // 1. 创建工作簿
        Workbook workbook = new XSSFWorkbook();
        // 2. 创建工作表
        Sheet sheet = workbook.createSheet();
        // 3. 创建行
        Row row = null;
        Cell cell;
        for (int rowNum = 0; rowNum < 10000000; rowNum++) {
            row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 2000; cellNum++) {
                cell = row.createCell(cellNum);
                cell.setCellValue(String.format("%s行%s列",rowNum,cellNum));
            }
        }
        FileOutputStream stream = new FileOutputStream("F:\\CodeSpace\\excel-demo\\07big.xls");
        workbook.write(stream);
        stream.close();
        System.out.println(System.currentTimeMillis() - start);
    }

    /**
     * SXSSFWorkbook 不会内存溢出，速度更快
     *
     * @throws Exception
     */
    @Test
    public void bigExcel07Super()throws Exception{
        long start = System.currentTimeMillis();
        // 1. 创建工作簿
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        // 2. 创建工作表
        Sheet sheet = workbook.createSheet();
        // 3. 创建行
        Row row = null;
        Cell cell;
        for (int rowNum = 0; rowNum < 10000; rowNum++) {
            row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 200; cellNum++) {
                cell = row.createCell(cellNum);
                cell.setCellValue(String.format("%s行%s列",rowNum,cellNum));
            }
        }
        FileOutputStream stream = new FileOutputStream("F:\\CodeSpace\\excel-demo\\07bigSuper.xls");
        workbook.write(stream);
        // 清除临时文件
        workbook.dispose();
        stream.close();
        System.out.println(System.currentTimeMillis() - start);
    }
}
