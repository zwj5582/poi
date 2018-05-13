/*
 *  Created by ZhongWenjie on 2018-05-07 13:38
 */

package com.excel.poi.annotation;

import com.excel.poi.support.CONTENT_TYPE;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.*;

@Target({ElementType.TYPE,ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface CellStyle {

    CONTENT_TYPE Type();

    BorderStyle BorderTop() default BorderStyle.THIN;

    BorderStyle BorderBottom() default BorderStyle.THIN;

    BorderStyle BorderLeft() default BorderStyle.THIN;

    BorderStyle BorderRight() default BorderStyle.THIN;

    HorizontalAlignment Alignment() default HorizontalAlignment.CENTER;

    VerticalAlignment VerticalAlignment() default VerticalAlignment.CENTER;

    HSSFColor.HSSFColorPredefined TopBorderColor() default HSSFColor.HSSFColorPredefined.BLACK;

    HSSFColor.HSSFColorPredefined BottomBorderColor() default HSSFColor.HSSFColorPredefined.BLACK;

    HSSFColor.HSSFColorPredefined LeftBorderColor() default HSSFColor.HSSFColorPredefined.BLACK;

    HSSFColor.HSSFColorPredefined RightBorderColor() default HSSFColor.HSSFColorPredefined.BLACK;

    HSSFColor.HSSFColorPredefined FillBackgroundColor() default HSSFColor.HSSFColorPredefined.WHITE;

    String DataFormat() default "";

    boolean ForceCover() default false;

}
