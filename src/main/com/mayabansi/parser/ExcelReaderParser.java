package com.mayabansi.parser;

import com.mayabansi.annotations.ExcelReader;
import com.mayabansi.annotations.TestCase;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;


/**
 * Created by IntelliJ IDEA.
 * User: Ravi Hasija
 * Date: Aug 10, 2011
 * Time: 7:30:02 PM
 * To change this template use File | Settings | File Templates.
 */
public class ExcelReaderParser {

    public String getFileName(Class<?> clazz) {
        ExcelReader excelReader = clazz.getAnnotation(ExcelReader.class);
        if (excelReader == null || excelReader.fileName().trim().length() == 0) {
            return "";
        }

        return excelReader.fileName();
    }

    public Sheet getSheet(String fileName) throws IOException, InvalidFormatException {
        InputStream inp = new FileInputStream(fileName);

        Workbook wb = WorkbookFactory.create(inp);
        Sheet sheet = wb.getSheetAt(0);
        return sheet;
    }

    public boolean doesClassHaveExcelReaderAnnotation(Class<?> clazz) {
        ExcelReader excelReader = clazz.getAnnotation(ExcelReader.class);
        if (excelReader == null) {
            return false;
        }
        return true;
    }

    public int getHeaderRow(Field field) {
        return field.getAnnotation(TestCase.class).headerRow();
    }

    public int getStartRow(Field field) {
        return field.getAnnotation(TestCase.class).startRow();
    }

    public int getEndColumn(Field field) {
        return field.getAnnotation(TestCase.class).endColumn();
    }

    public List<String> getHeaderNames(final Sheet sheet, final int headerPos, final int endColumn) {

        Row row = sheet.getRow(headerPos-1);

        final List<String> headerNames = new ArrayList<String>();
        for (int i = 0; i < endColumn; i++) {
            Cell cell = row.getCell(i);
            if (cell != null) {
                String header = cell.getStringCellValue();
                header = header.replaceAll(" ", "");
                headerNames.add(header);
            }
        }
        return headerNames;
    }

    public boolean isBlank(final Sheet sheet, final int rowPos, final int endColumn)
    {
        Row row = sheet.getRow(rowPos-1);

        if (row == null) {
            return true;
        }
        
        for (int i = 0; i < endColumn; i++) {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
                return false;
            }
        }
        return true;
    }

    public Object getDataFromCell(final Cell cell) {

        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            return cell.getNumericCellValue();
        } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
            return cell.getStringCellValue();
        }

        return null;
    }

    public void parseData(
            final Sheet sheet,
            final int startRow,
            final int endColumn,
            final List<String> propertyNames,
            final Class clazzOfTestCase) {

        int rowCounter = startRow;

        while (!isBlank(sheet, rowCounter, endColumn)) {
            Row row = sheet.getRow(rowCounter - 1);
            for (int i = 0; i < endColumn; i++) {
                Cell cell = row.getCell(i);
                if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
                    try {
                        Method method = clazzOfTestCase.getMethod("set"+ propertyNames.get(i), Object.class);
                    } catch (NoSuchMethodException e) {
                        e.printStackTrace();  //To change body of catch statement use File | Settings | File Templates.
                    }
                }
            }
        }





    }

    public int parse(Class<?> clazz) throws Exception {

        if (!doesClassHaveExcelReaderAnnotation(clazz)) {
            return 0;
        }

        String fileName = getFileName(clazz);

        Sheet sheet = getSheet(fileName);

        Field[] fields = clazz.getDeclaredFields();
        int count = 0;


        for (Field field : fields) {
            if (field.isAnnotationPresent(TestCase.class)) {
                try {
                    //method.invoke(null);
                    System.out.println("Field name: " + field.getName());
                    count++;
                } catch (Exception e) {
                }
            }
        }

        return count;
    }
}
