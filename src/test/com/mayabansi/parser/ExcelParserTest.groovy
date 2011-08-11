package com.mayabansi.parser

import spock.lang.Specification
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Cell

/**
 * Created by IntelliJ IDEA.
 * User: Ravi Hasija
 * Date: Aug 10, 2011
 * Time: 7:58:04 PM
 * To change this template use File | Settings | File Templates.
 */
class ExcelParserTest extends Specification {

    def "when no method has TestCase annotation then returns 0 count"() {
        when:
            ExcelReaderParser excelParser = new ExcelReaderParser()
            
        then:
            excelParser.parse(ClassWithNoTestCaseAnnotation.class) == 0
    }

    def "when method has TestCase annotation then returns correct count"() {
        when:
            ExcelReaderParser excelParser = new ExcelReaderParser()

        then:
            excelParser.doesClassHaveExcelReaderAnnotation(ClassWithTestCaseAnnotation.class) == true
    }

     def "get correct file name from annotation"() {
        when:
            ExcelReaderParser excelParser = new ExcelReaderParser()

        then:
            excelParser.getFileName(ClassWithTestCaseAnnotation.ClassWithExcelReaderAnnotationAndHasFileNameThatDoesExist.class).contains("someTest.xls")
    }

    def "file does not exist"() {
        when:
            ExcelReaderParser excelParser = new ExcelReaderParser()
            def fileName = excelParser.getFileName(ClassWithTestCaseAnnotation.ClassWithExcelReaderAnnotationAndHasFileNameThatDoesNotExist.class)
            excelParser.getSheet(fileName)

        then:
            thrown (IOException)     
    }

    def "file does exist"() {
        when:
            ExcelReaderParser excelParser = new ExcelReaderParser()
            def fileName = excelParser.getFileName(ClassWithTestCaseAnnotation.ClassWithExcelReaderAnnotationAndHasFileNameThatDoesExist.class)
            def sheet = excelParser.getSheet(fileName)

        then:
            sheet != null
    }

    def "verify header row, start row, end column"() {
        when:
            ExcelReaderParser excelParser = new ExcelReaderParser()
            def headerRow = excelParser.getHeaderRow(ClassWithTestCaseAnnotation.ClassWithExcelReaderAnnotationAndHasFileNameThatDoesExist.class.getDeclaredFields()[0])
            def startRow = excelParser.getStartRow(ClassWithTestCaseAnnotation.ClassWithExcelReaderAnnotationAndHasFileNameThatDoesExist.class.getDeclaredFields()[0])
            def endColumn = excelParser.getEndColumn(ClassWithTestCaseAnnotation.ClassWithExcelReaderAnnotationAndHasFileNameThatDoesExist.class.getDeclaredFields()[0])

        then:
            headerRow == 1
            startRow == 2
            endColumn == 4
    }

    def "verify header names are correct"() {
        when:
            ExcelReaderParser excelParser = new ExcelReaderParser()
            def fileName = excelParser.getFileName(ClassWithTestCaseAnnotation.ClassWithExcelReaderAnnotationAndHasFileNameThatDoesExist.class)
            def sheet = excelParser.getSheet(fileName)
            List<String> headerNames = excelParser.getHeaderNames(sheet, 1, 4)
        then:
            headerNames.size() == 4
            headerNames[0] == "Id"
            headerNames[1] == "AccountNumber"
            headerNames[2] == "CustomerNumber"
            headerNames[3] == "Balance"
    }

    def "return blank returns true for blank row"() {
        when:
            ExcelReaderParser excelParser = new ExcelReaderParser()
            def fileName = excelParser.getFileName(ClassWithTestCaseAnnotation.ClassWithExcelReaderAnnotationAndHasFileNameThatDoesExist.class)
            def sheet = excelParser.getSheet(fileName)
        then:
            excelParser.isBlank(sheet, 10, 4) == true
    }

    def "return blank returns for non-blank row"() {
        when:
            ExcelReaderParser excelParser = new ExcelReaderParser()
            def fileName = excelParser.getFileName(ClassWithTestCaseAnnotation.ClassWithExcelReaderAnnotationAndHasFileNameThatDoesExist.class)
            def sheet = excelParser.getSheet(fileName)
        then:
            excelParser.isBlank(sheet, 1, 4) == false
    }

    def "data from cell (1,1) returns 1"(def rowPos, def cellPos, def expectedResult) {
        
            ExcelReaderParser excelParser = new ExcelReaderParser()
            def fileName = excelParser.getFileName(ClassWithTestCaseAnnotation.ClassWithExcelReaderAnnotationAndHasFileNameThatDoesExist.class)
            def sheet = excelParser.getSheet(fileName)

        expect:
            Cell cell = sheet.getRow(rowPos-1).getCell(cellPos-1)
            excelParser.getDataFromCell(cell) == expectedResult

        where:
            rowPos | cellPos | expectedResult
            2   | 1          | 1
            2   | 2          | "a11"
            2   | 3          | "customer1"
            2   | 4          | 100
    }
}
