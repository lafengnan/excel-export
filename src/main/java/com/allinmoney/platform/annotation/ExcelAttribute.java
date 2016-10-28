package com.allinmoney.platform.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Created by chris on 16/4/27.
 */

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelAttribute {

    // column name in excel
    String title();

    // view group
    Class<?>[] groups() default {};

    // key-value map
    Translate[] translate() default {};

    // column mark A,B,C,D...
    String column() default "";

    // prompt info
    String prompt() default "";

    // date format
    String format() default "yyyy-MM-dd HH:mm:ss";

    // only selective columns
    String[] combo() default {};

    // export or not export
    boolean isExport() default true;

    // important fields
    boolean isMark() default false;

    // summarize current column?
    boolean isSum() default false;

}
