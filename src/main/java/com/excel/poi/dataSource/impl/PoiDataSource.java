/*
 *  Created by ZhongWenjie on 2018-05-08 12:57
 */

package com.excel.poi.dataSource.impl;

import com.excel.poi.support.MappingObjectProxy;
import com.excel.poi.dataSource.DataSource;

public abstract class PoiDataSource<T> implements DataSource<T,MappingObjectProxy> {

    protected Iterable<MappingObjectProxy> dataSources;

    protected T dataSource;

    protected int length = 0;

    public PoiDataSource(T dataSource) {
        this.dataSource = dataSource;
        initPoiDataSource();
    }

    protected PoiDataSource(Iterable<MappingObjectProxy> dataSources,int length){
        this.dataSources = dataSources;
        this.length = length;
    }

    protected abstract void initPoiDataSource();

    public abstract int length();

}
