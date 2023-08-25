package com.ck.utils.excel;

import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.*;

public class ExcelTemplateHelper {

    public static Map<String, Object> readData(XSSFSheet sheet, String loopKey, XSSFSheet srcSheet) {
        return readDataFromTpl(sheet, loopKey, readTemplate(srcSheet, loopKey));
    }

    private static Map<String, Object> readDataFromTpl(XSSFSheet sheet, String loopKey, Map<Integer, Map<Boolean, Map<Integer, String>>> tpl) {
        Map<String, Object> result = new HashMap<>();
        List<Map<String, Object>> loop = new ArrayList<>();
        boolean startLoop = false;
        int startLoopRow = 0;
        for (int r = sheet.getFirstRowNum(); r <= sheet.getLastRowNum(); r++) {
            if (!tpl.containsKey(r) && !startLoop) {
                continue;
            }
            XSSFRow row = sheet.getRow(r);
            if (tpl.containsKey(r)) {
                if (tpl.get(r).containsKey(false)) {
                    for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
                        if (!tpl.get(r).get(false).containsKey(c)) {
                            continue;
                        }
                        XSSFCell cell = row.getCell(c);
                        String value = cell.getStringCellValue();
                        result.put(tpl.get(r).get(false).get(c), value);
                    }
                } else {
                    startLoop = true;
                    startLoopRow = r;
                    Map<String, Object> temp = new HashMap<>();
                    loop.add(temp);
                    for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
                        if (!tpl.get(r).get(true).containsKey(c)) {
                            continue;
                        }
                        XSSFCell cell = row.getCell(c);
                        String value = cell.getStringCellValue();
                        temp.put(tpl.get(r).get(true).get(c), value);
                    }
                }
            } else {
                Map<String, Object> temp = new HashMap<>();
                loop.add(temp);
                for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
                    if (!tpl.get(startLoopRow).get(true).containsKey(c)) {
                        continue;
                    }
                    XSSFCell cell = row.getCell(c);
                    String value = cell.getStringCellValue();
                    temp.put(tpl.get(startLoopRow).get(true).get(c), value);
                }
            }
        }
        result.put(loopKey, loop);
        return result;
    }

    private static Map<Integer, Map<Boolean, Map<Integer, String>>> readTemplate(XSSFSheet sheet, String loopKey) {
        Map<Integer, Map<Boolean, Map<Integer, String>>> result = new LinkedHashMap<>();
        for (int r = sheet.getFirstRowNum(); r <= sheet.getLastRowNum(); r++) {
            XSSFRow row = sheet.getRow(r);
            for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
                XSSFCell cell = row.getCell(c);
                String value = cell.getStringCellValue();
                if (StringUtils.isNotBlank(value)) {
                    if (value.startsWith("{{") && value.endsWith("}}")) {
                        if (value.startsWith("{{" + loopKey + ".")) {
                            result.computeIfAbsent(r, key -> new HashMap<>()).computeIfAbsent(true, key -> new HashMap<>()).put(c, value.substring(3 + loopKey.length(), value.length() - 2));
                        } else {
                            result.computeIfAbsent(r, key -> new HashMap<>()).computeIfAbsent(false, key -> new HashMap<>()).put(c, value.substring(2, value.length() - 2));
                        }
                    }
                }
            }
        }
        return result;
    }

    public static void handleSheet(XSSFSheet sheet, Map<String, String> staticSource, List<Map<String, Object>> dynamicSourceList) {
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            XSSFRow row = sheet.getRow(i);
            Map<String, Object> dynamicSource = parseDynamicRow(row, dynamicSourceList);
            if (dynamicSource != null) {
                i = handleDynamicRows(dynamicSource, sheet, i);
            } else {
                replaceRowValue(row, staticSource, null);
            }
        }
    }

    private static Map<String, Object> parseDynamicRow(XSSFRow row, List<Map<String, Object>> dynamicSourceList) {
        if (dynamicSourceList.isEmpty()) {
            return null;
        }
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            XSSFCell cell = row.getCell(i);
            String value = cell.getStringCellValue();
            if (value != null) {
                for (Map<String, Object> current : dynamicSourceList) {
                    String id = MapUtils.getString(current, "loopId");
                    if (value.startsWith("{{" + id + ".")) {
                        return current;
                    }
                }
            }
        }
        return null;
    }

    private static int handleDynamicRows(Map<String, Object> dynamicSource, XSSFSheet sheet, int rowIndex) {
        if (dynamicSource.isEmpty()) {
            return rowIndex;
        }
        String id = MapUtils.getString(dynamicSource, "loopId");
        List<Map<String, String>> dataList = (List<Map<String, String>>) dynamicSource.get("dataList");
        if (dataList == null) {
            return rowIndex;
        }
        int rows = dataList.size();
        // 因为模板行本身占1行，所以-1
        int copyRows = rows - 1;
        if (copyRows > 0) {
            // shiftRows: 从startRow到最后一行，全部向下移copyRows行
            sheet.shiftRows(rowIndex, sheet.getLastRowNum(), copyRows, true, false);
            // 拷贝策略
            CellCopyPolicy cellCopyPolicy = new CellCopyPolicy();
            cellCopyPolicy.setCopyCellValue(true);
            cellCopyPolicy.setCopyCellStyle(true);
            // 这里模板row已经变成了startRow + copyRows,
            int templateRow = rowIndex + copyRows;
            // 因为下移了，所以要把模板row拷贝到所有空行
            for (int i = 0; i < copyRows; i++) {
                sheet.copyRows(templateRow, templateRow, rowIndex + i, cellCopyPolicy);
            }
        }
        // 替换动态行的值
        for (int j = rowIndex; j < rowIndex + rows; j++) {
            replaceRowValue(sheet.getRow(j), dataList.get(j - rowIndex), id);
        }
        return rowIndex + copyRows;
    }

    private static void replaceRowValue(XSSFRow row, Map<String, String> map, String prefixKey) {
        if (map.isEmpty()) {
            return;
        }
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            XSSFCell cell = row.getCell(i);
            replaceCellValue(cell, map, prefixKey);
        }
    }

    private static void replaceCellValue(XSSFCell cell, Map<String, String> map, String prefixKey) {
        if (cell == null) {
            return;
        }
        String cellValue = cell.getStringCellValue();
        if (StringUtils.isBlank(cellValue)) {
            return;
        }
        boolean flag = false;
        prefixKey = StringUtils.isBlank(prefixKey) ? "" : (prefixKey + ".");
        for (Map.Entry<String, String> current : map.entrySet()) {
            // 循环所有，因为可能一行有多个占位符
            String template = "{{" + prefixKey + current.getKey() + "}}";
            if (cellValue.contains(template)) {
                String value = current.getValue();
                if (value == null) {
                    value = "";
                }
                cellValue = cellValue.replace(template, value);
                flag = true;
            }
        }
        if (flag) {
            cell.setCellValue(cellValue);
        }
    }
}
