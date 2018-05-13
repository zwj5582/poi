/*
 *  Created by ZhongWenjie on 2018-05-07 17:32
 */

package com.excel.poi.reader;

import com.excel.poi.context.PoiContext;

public interface PoiPropertiesReader {

    PoiContext readProperties(Class<?> entityClass) throws InstantiationException, IllegalAccessException, ClassNotFoundException;

}
