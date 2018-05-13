/*
 *  Created by ZhongWenjie on 2018-05-07 11:12
 */

package com.excel.poi.annotation;

import java.lang.annotation.*;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface DisplayName {

    String value() default "";

}
