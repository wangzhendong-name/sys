package com.ese.sys.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelUtils {
    public static List<Map<String, Object>> readExcel(String fileName, InputStream inputStream) throws IOException {

        List<Map<String, Object>> list = null;

        boolean ret = isXls(fileName);

        Workbook workbook = null;

        if (ret == true) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            workbook = new XSSFWorkbook(inputStream);
        }

        Sheet sheet = workbook.getSheetAt(0);
        // 标题
        Row row = sheet.getRow(0);

        // 获取row里面最后一个单元格的编号
        short lastCellNum = row.getLastCellNum();
        // 获取最后一行的行号
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 1; i <= lastRowNum; i++) {
            HashMap<String, Object> map = new HashMap<>();
            Row row1 = sheet.getRow(i);

            for (int j = 0; j < lastCellNum; j++) {
                Cell cell = row1.getCell(j);
                if (cell == null) {
                    continue;
                }
                cell.setCellType(CellType.STRING);
                map.put(row.getCell(j).getStringCellValue(), cell.getStringCellValue());

                System.out.println(cell.getStringCellValue());
            }
            list.add(map);

        }
        return list;
    }

    /**
     * 判断excel文件的后缀，如果是xls后缀，返回true，是xlsx后缀，返回false
     * @param filename
     * @return
     */
    private static boolean isXls(String filename) {

        if (filename.matches("^.+\\.(?i)(xls$)")) {
            return true;
        } else {
            return false;
        }

    }
}
