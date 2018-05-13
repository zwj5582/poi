/*
 *  Created by ZhongWenjie on 2018-05-10 22:37
 */

package com.excel.poi.annotation;

import java.lang.annotation.*;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Condition {

    String value();

    ConditionalStyle True();

    ConditionalStyle False();

}
