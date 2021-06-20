package com.demo.read;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.Iterator;

public class ExcelRead {


    static final String PATH = "F:\\CodeSpace\\excel-demo\\";

    DecimalFormat df = new DecimalFormat("0");

    /**
     * xls 文件读取 使用 HSSFWorkbook
     *
     * @throws Exception
     */
    @Test
    public void testRead03() throws Exception {
        // 1. 获取文件流
        FileInputStream stream = new FileInputStream(PATH + "03.xls");
        // 2. 创建工作簿
        Workbook workbook = new HSSFWorkbook(stream);
        // 3. 获取页
        Sheet sheet = workbook.getSheetAt(0);
        // 4. 获取行
        Row row = sheet.getRow(0);
        // 获取单元格
        Cell cell = row.getCell(0);
        stream.close();
        System.out.println(cell);
    }

    /**
     * xlsx 文件读取使用 XSSFWorkbook
     *
     * @throws Exception
     */
    @Test
    public void testRead07() throws Exception {
        // 1. 获取文件流
        FileInputStream stream = new FileInputStream(PATH + "07.xlsx");
        // 2. 创建工作簿
        Workbook workbook = new XSSFWorkbook(stream);
        // 3. 获取页
        Sheet sheet = workbook.getSheetAt(0);
        // 4. 获取行
        Row row = sheet.getRow(0);
        // 获取单元格
        Cell cell = row.getCell(0);
        stream.close();
        System.out.println(cell);
    }

    /**
     * 读取 不同格式单元格
     * @throws Exception
     */
    @Test
    public void testCellType() throws Exception {
        // 1. 获取文件流
        FileInputStream stream = new FileInputStream(PATH + "07big.xlsx");
        // 2. 创建工作簿
        XSSFWorkbook workbook = new XSSFWorkbook(stream);
        // 3. 获取所有 Sheet 页
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        Sheet sheet;
        Row row;
        Cell cell;
        Iterator<Row> rowIterator;
        Iterator<Cell> cellIterator;
        // 4. 获取计算公式
        XSSFFormulaEvaluator evaluator = new XSSFFormulaEvaluator(workbook);
        // 5. 遍历 Sheet 页
        while (sheetIterator.hasNext()) {
            sheet = sheetIterator.next();
            // 6. 获取所有行
            rowIterator = sheet.rowIterator();
            // 7. 遍历 所有行
            while (rowIterator.hasNext()) {
                row = rowIterator.next();
                // 8. 获取当前行单元格
                cellIterator = row.cellIterator();
                // 9. 遍历单元格
                while (cellIterator.hasNext()) {
                    cell = cellIterator.next();
                    if (cell != null) {
                        // 10 .获取单元格类型
                        CellType cellType = cell.getCellType();
                        if (CellType.BOOLEAN.compareTo(cellType) == 0) {
                            System.out.println(cell.getBooleanCellValue());
                        } else if (CellType.NUMERIC.compareTo(cellType) == 0) {
                            if (DateUtil.isCellDateFormatted(cell)) {
                                // 日期格式处理
                                System.out.println(cell.getDateCellValue());
                            } else {
                                // 数字格式
                                System.out.println(new BigDecimal(cell.getNumericCellValue()));
                            }
                        } else if (CellType.STRING.compareTo(cellType) == 0) {
                            System.out.println(cell.getStringCellValue());
                        } else if (CellType.ERROR.compareTo(cellType) == 0) {
                            System.out.println("类型错误");
                        }else if (CellType.FORMULA.compareTo(cellType) ==0){
                            // 数学公式处理
                            System.out.println(cell.getCellFormula());
                            CellValue evaluate = evaluator.evaluate(cell);
                            System.out.println(evaluate.formatAsString());
                        }
                    }
                }
            }
        }
        stream.close();
    }

}
