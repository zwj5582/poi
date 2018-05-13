/*
 *  Created by ZhongWenjie on 2018-05-07 17:45
 */

package com.excel.poi.dataSource;

public interface DataSource<T,E> {

    Iterable<E> fetchDataSource();

}
