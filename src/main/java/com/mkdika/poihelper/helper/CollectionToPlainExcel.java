package com.mkdika.poihelper.helper;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.*;

/**
 * Simple Java Collections to Plain (without style) Excel File Generator.
 *
 * <p>
 * This class only supported for simple fields Type, like:
 * String, Date, Double, Integer, Boolean.
 * </p>
 *
 * @author Maikel Chandika </mkdika@gmail.com>
 * @version 1.0.0
 * @since 11th Oct 2018
 */
public class CollectionToPlainExcel {

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private CellStyle cellStyle;
    private String sheetName;
    private Collection data;
    private boolean showFieldName;

    private static Map<String, Integer> basicTypeMap;

    static {
        basicTypeMap = new HashMap<>();
        basicTypeMap.put("java.util.Date", 1);
        basicTypeMap.put("java.lang.Double", 2);
        basicTypeMap.put("java.lang.Float", 2);
        basicTypeMap.put("java.lang.Integer", 2);
        basicTypeMap.put("java.lang.Long", 2);
        basicTypeMap.put("java.lang.Short", 2);
        basicTypeMap.put("java.lang.Byte", 2);
        basicTypeMap.put("java.lang.Boolean", 3);
        basicTypeMap.put("java.lang.String", 4);
    }

    public CollectionToPlainExcel() {
        this.sheetName = "Sheet1";
        this.showFieldName = true;
    }

    public void writeToExcel(String fileName) throws IOException, ClassNotFoundException, IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        if (validateAndBuild()) {
            FileOutputStream fos = new FileOutputStream(fileName);
            writeToExcel(fos);
        }
    }

    private void writeToExcel(FileOutputStream fos) throws IOException {
        workbook.write(fos);
        workbook.close();
    }

    public byte[] writeToExcel() throws IOException, ClassNotFoundException, IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        if (validateAndBuild()) {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            workbook.write(bos);
            workbook.close();
            bos.close();
            return bos.toByteArray();
        } else {
            return new byte[]{};
        }
    }

    private boolean validateAndBuild() throws ClassNotFoundException, IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        if (data == null || data.size() <= 0) return false;

        String className = data.iterator().next().getClass().getCanonicalName();
        Field[] fields = Class.forName(className).getDeclaredFields();

        this.workbook = new XSSFWorkbook();
        this.sheet = workbook.createSheet(sheetName);

        // Excel date field formatter
        cellStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-mm-yy hh:mm:ss"));

        int rowNum = 0;
        int[] types;
        boolean basicType = basicTypeMap.containsKey(className);

        if (basicType) {
            types = new int[]{basicTypeMap.get(className)};
        } else {
            types = new int[fields.length];
        }

        if (showFieldName && fields != null && fields.length > 0) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;

            if (basicType) {
                Cell cell = row.createCell(colNum++);
                cell.setCellValue("data");
            } else {
                for (Field f : fields) {
                    if (basicTypeMap.containsKey(f.getType().getCanonicalName())) {
                        types[colNum] = basicTypeMap.get(f.getType().getCanonicalName());
                    } else {
                        throw new IllegalArgumentException("The data fields type should be only in " +
                                "[Date, Double, Float, Integer, Long, Short, Byte, Boolean, String], " +
                                "but found: " + f.getType().getCanonicalName());
                    }
                    Cell cell = row.createCell(colNum++);
                    cell.setCellValue(f.getName());
                }
            }
        }

        Iterator it = data.iterator();
        while (it.hasNext()) {
            Object obj = it.next();
            if (obj != null) {
                Row row = sheet.createRow(rowNum++);
                int colNum = 0;

                if (basicType) {
                    Cell cell = row.createCell(colNum++);
                    insertCell(cell, types[colNum - 1], obj);
                } else {
                    for (Field f : fields) {
                        Cell cell = row.createCell(colNum++);
                        Object result = PropertyUtils.getSimpleProperty(obj, f.getName());
                        if (result != null) {
                            insertCell(cell, types[colNum - 1], result);
                        }
                    }
                }
            }
        }
        return true;
    }

    /*
    1. Date
    2. Double,Float, Integer, Long, Byte, Short
    3. Boolean
    4. String or others
    */
    private void insertCell(Cell cell, int type, Object obj) {
        switch (type) {
            case 1:
                cell.setCellValue((Date) obj);
                cell.setCellStyle(cellStyle);
                break;
            case 2:
                cell.setCellValue(Double.valueOf(obj.toString()));
                break;
            case 3:
                cell.setCellValue((Boolean) obj);
                break;
            default:
                cell.setCellValue(obj.toString());
        }
    }

    // Setters & Getters
    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public Collection getData() {
        return data;
    }

    public void setData(Collection data) {
        this.data = data;
    }

    public boolean isShowFieldName() {
        return showFieldName;
    }

    public void setShowFieldName(boolean showFieldName) {
        this.showFieldName = showFieldName;
    }
}
