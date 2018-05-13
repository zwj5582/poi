/*
 *  Created by ZhongWenjie on 2018-05-07 21:44
 */

package com.excel.poi.dataSource.impl;

import com.excel.poi.support.MappingObjectProxy;

public class ComposeDataSource extends PoiDataSource<Object> {

    public ComposeDataSource(Object dataSource) {
        super(dataSource);
    }

    protected ComposeDataSource(Iterable<MappingObjectProxy> dataSources, int length) {
        super(dataSources, length);
    }

    @Override
    protected void initPoiDataSource() {
        if (dataSource instanceof Iterable ) {
            this.dataSources = new IterableDataSource((Iterable) dataSource)
                    .fetchDataSource();
        } else{
            this.dataSources = new ArrayDataSource((Object[]) dataSource)
                    .fetchDataSource();
        }
    }

    @Override
    public Iterable<MappingObjectProxy> fetchDataSource() {
        return super.dataSources;
    }

    @Override
    public int length() {
        return super.length;
    }
}
