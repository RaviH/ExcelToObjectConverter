package com.mayabansi.parser;

import com.mayabansi.annotations.ExcelReader;
import com.mayabansi.annotations.TestCase;

/**
 * Created by IntelliJ IDEA.
 * User: Ravi Hasija
 * Date: Aug 10, 2011
 * Time: 8:08:10 PM
 * To change this template use File | Settings | File Templates.
 */
@ExcelReader
public class ClassWithTestCaseAnnotation {

    @TestCase Integer foo;

    @ExcelReader(fileName = "some.xlsx")
    public class ClassWithExcelReaderAnnotationAndHasFileNameThatDoesNotExist {

    }

    @ExcelReader(fileName = "D:\\code\\IProjects\\Excel\\test\\com\\mayabansi\\parser\\excelfiles\\someTest.xls")
    public class ClassWithExcelReaderAnnotationAndHasFileNameThatDoesExist {

        @TestCase(startRow = 2, headerRow = 1, endColumn = 4) Customer customer;
    }

    public class Customer {
        private String customerNumber;

        private String accountNumber;

        private Double amount;

        private Double id;

        public String getCustomerNumber() {
            return customerNumber;
        }

        public void setCustomerNumber(String customerNumber) {
            this.customerNumber = customerNumber;
        }

        public String getAccountNumber() {
            return accountNumber;
        }

        public void setAccountNumber(String accountNumber) {
            this.accountNumber = accountNumber;
        }

        public Double getAmount() {
            return amount;
        }

        public void setAmount(Double amount) {
            this.amount = amount;
        }

        public Double getId() {
            return id;
        }

        public void setId(Double id) {
            this.id = id;
        }
    }
}
