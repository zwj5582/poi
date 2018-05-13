/*
 *  Created by ZhongWenjie on 2018-05-07 14:12
 */

package com.excel.poi.annotation;

import org.apache.poi.hssf.util.HSSFColor;

import java.lang.annotation.*;

@Target({ElementType.TYPE,ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface HeaderFont {

    String FontName() default "宋体";

    boolean Bold() default true;

    HSSFColor.HSSFColorPredefined Color() default HSSFColor.HSSFColorPredefined.WHITE;

    short FontHeightInPoints() default 12;

    boolean ForceCover() default false;

}
