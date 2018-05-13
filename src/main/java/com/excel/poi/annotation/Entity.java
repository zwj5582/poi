/*
 *  Created by ZhongWenjie on 2018-05-07 10:49
 */

package com.excel.poi.annotation;

import com.excel.poi.support.CONTENT_TYPE;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.*;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Entity {

    String title();

    int columnWidth() default 18 * 256;

    int columnHeight() default 30;

    int titleColumnHeight() default 40;

    CellStyle titleStyle() default
            @CellStyle(
                Type = CONTENT_TYPE.TITLE,
                BorderTop = BorderStyle.THICK,
                BorderBottom = BorderStyle.THICK,
                BorderLeft = BorderStyle.THICK,
                BorderRight = BorderStyle.THICK,
                Alignment = HorizontalAlignment.CENTER,
                VerticalAlignment = VerticalAlignment.CENTER,
                TopBorderColor = HSSFColor.HSSFColorPredefined.BLACK,
                BottomBorderColor = HSSFColor.HSSFColorPredefined.BLACK,
                LeftBorderColor = HSSFColor.HSSFColorPredefined.BLACK,
                RightBorderColor = HSSFColor.HSSFColorPredefined.BLACK,
                FillBackgroundColor = HSSFColor.HSSFColorPredefined.GREY_25_PERCENT,
                DataFormat = "",
                ForceCover = false
            );

    Font titleFont() default
            @Font(
                    FontName = "黑体",
                    Bold = true,
                    FontHeightInPoints = 24,
                    Color = HSSFColor.HSSFColorPredefined.WHITE
            );

}
