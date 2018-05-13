/*
 *  Created by ZhongWenjie on 2018-05-07 12:54
 */

package com.excel.poi.annotation;

import org.apache.poi.ss.usermodel.CellType;

import java.lang.annotation.*;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Type {

    CellType value() default CellType.STRING;

}
