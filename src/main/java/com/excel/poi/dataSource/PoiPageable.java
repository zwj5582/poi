/*
 *  Created by ZhongWenjie on 2018-05-08 15:10
 */

package com.excel.poi.dataSource;

import com.excel.poi.exception.TooManyDataExportException;
import com.excel.poi.dataSource.impl.PoiDataSource;

public interface PoiPageable extends Pageable {

    Iterable<PoiDataSource> transformDataSource() throws TooManyDataExportException;

    short size();

    short currentPage();

}
