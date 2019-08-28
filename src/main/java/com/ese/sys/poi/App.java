package com.ese.sys.poi;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class App {

    public static void readXLSX() throws IOException {

        List<Map<String, Object>> list = null;
        // 读取文件
        FileInputStream fileInputStream = new FileInputStream("d:/123.xlsx");
        // 创建工作簿对象
        XSSFWorkbook sheets = new XSSFWorkbook(fileInputStream);
        // 根据索引获取Sheet对象
        XSSFSheet sheetAt = sheets.getSheetAt(0);
        // 获取Row对象
        XSSFRow row = sheetAt.getRow(0);
        // 获取Row对象
        XSSFCell cell = row.getCell(0);
        // 获取d单元的文本内容
        String stringCellValue = cell.getStringCellValue();
        System.out.println(stringCellValue);

        XSSFCell cell1 = row.getCell(1);
        // 数字格式的单元格
        cell1.setCellType(CellType.STRING);
        System.out.println(cell1.getStringCellValue());


        sheets.close();
        fileInputStream.close();
    }

    public static void writeXLSX() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(0);
        cell.setCellValue("zhangsan");

        // 生成excel时，需要输出流的对象
        FileOutputStream fileOutputStream = new FileOutputStream("d:/b.xlsx");
        workbook.write(fileOutputStream);
    }


    public static void main(String[] args) {
        try {
            readXLSX();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
