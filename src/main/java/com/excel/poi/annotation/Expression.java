/*
 *  Created by ZhongWenjie on 2018-05-07 22:01
 */

package com.excel.poi.annotation;

import java.lang.annotation.*;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Expression {

    String value();

}
