/*
 *  Created by ZhongWenjie on 2018-05-10 22:09
 */

package com.excel.poi.annotation;


import java.lang.annotation.*;

@Target({ElementType.TYPE,ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ConditionalStyle {

    String BorderTop() default "";

    String BorderBottom() default "";

    String BorderLeft() default "";

    String BorderRight() default "";

    String Alignment() default "";

    String VerticalAlignment() default "";

    String TopBorderColor() default "";

    String BottomBorderColor() default "";

    String LeftBorderColor() default "";

    String RightBorderColor() default "";

    String FillBackgroundColor() default "";

    String DataFormat() default "";


    String FontName() default "";

    String Bold() default "";

    String Color() default "";

    String FontHeightInPoints() default "";

}
