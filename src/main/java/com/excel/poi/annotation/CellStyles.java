/*
 *  Created by ZhongWenjie on 2018-05-10 14:45
 */

package com.excel.poi.annotation;

import java.lang.annotation.*;

@Target({ElementType.TYPE,ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface CellStyles {

    CellStyle[] value();

}
