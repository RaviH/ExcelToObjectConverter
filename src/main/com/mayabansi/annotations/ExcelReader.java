package com.mayabansi.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Created by IntelliJ IDEA.
 * User: Ravi Hasija
 * Date: Aug 10, 2011
 * Time: 7:23:27 PM
 * To change this template use File | Settings | File Templates.
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelReader {
    String fileName() default "";
}
