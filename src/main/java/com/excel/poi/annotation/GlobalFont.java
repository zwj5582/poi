/*
 *  Created by ZhongWenjie on 2018-05-10 16:26
 */

package com.excel.poi.annotation;

import org.apache.poi.hssf.util.HSSFColor;

import java.lang.annotation.*;

@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface GlobalFont {

    String FontName() default "宋体";

    boolean Bold() default false;

    HSSFColor.HSSFColorPredefined Color() default HSSFColor.HSSFColorPredefined.BLACK;

    short FontHeightInPoints() default 12;

    boolean ForceCover() default false;

}
