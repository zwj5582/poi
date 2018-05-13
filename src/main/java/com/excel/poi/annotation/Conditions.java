/*
 *  Created by ZhongWenjie on 2018-05-11 10:43
 */

package com.excel.poi.annotation;


import java.lang.annotation.*;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Conditions {

    Condition[] values();

}
