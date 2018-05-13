/*
 *  Created by ZhongWenjie on 2018-05-12 15:36
 */

package com.excel.poi.annotation;

import java.lang.annotation.*;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface RowCondition {

    String value();

    ConditionalStyle True();

    ConditionalStyle False();

}
